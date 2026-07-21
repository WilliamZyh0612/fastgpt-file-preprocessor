from __future__ import annotations
import json, re, time, urllib.error, urllib.request
from dataclasses import dataclass
from typing import Callable
from .config import Settings

CATEGORIES = {"产品资料", "操作资料", "维修资料", "配件资料", "业务规则", "待人工确认"}
DOC_FIELDS = {"categories", "machine_models", "cnc_systems", "visibility", "version", "release_date", "summary", "knowledge_points", "keywords"}
STRUCTURED = {
 "product_records": {"product_model","product_name","processing_object","processing_range","key_parameters","accuracy","standard_configuration","optional_configuration","applicable_scenarios","limitations","cnc_system","evidence_refs","confidence"},
 "fault_records": {"machine_model","cnc_system","fault_symptom","alarm_code","operating_condition","possible_causes","troubleshooting_steps","solution","safety_warning","requires_engineer","evidence_refs","confidence"},
 "part_records": {"part_name","part_model","drawing_number","applicable_machine_models","installation_position","technical_specification","replacement_parts","replacement_limits","evidence_refs","confidence"},
 "business_rule_records": {"business_module","rule_name","applicable_roles","trigger_conditions","allowed_actions","prohibited_actions","approval_requirements","status_transition","exception_handling","evidence_refs","confidence"},
}
AI_SCHEMA = {"name":"knowledge_preprocessor_result","strict":True,"schema":{"type":"object","additionalProperties":False,"required":["document_metadata",*STRUCTURED],"properties":{}}}

class AIResponseError(ValueError): pass

def _string(value: object, label: str) -> None:
    if not isinstance(value, str) or len(value)>4000: raise AIResponseError(f"{label} 必须是长度受限字符串")
def _refs(value: object, ids: set[str] | None) -> None:
    if not isinstance(value,list) or not all(isinstance(x,str) for x in value): raise AIResponseError("evidence_refs 类型不合法")
    if ids is not None and not set(value).issubset(ids): raise AIResponseError("AI 返回不存在的 evidence_refs")

def validate_ai_result(value: object, evidence_ids: set[str] | None = None) -> dict:
    if not isinstance(value,dict) or set(value)!={"document_metadata",*STRUCTURED}: raise AIResponseError("AI 顶层字段不符合严格 Schema")
    meta=value["document_metadata"]
    if not isinstance(meta,dict) or set(meta)!=DOC_FIELDS: raise AIResponseError("document_metadata 字段不合法")
    for name in ("visibility","version","release_date","summary"):_string(meta[name],name)
    if meta["release_date"] not in {"待人工确认",""} and not re.fullmatch(r"\d{4}-\d{2}-\d{2}",meta["release_date"]): raise AIResponseError("日期必须为 yyyy-mm-dd")
    if meta["version"] not in {"待人工确认",""} and not re.fullmatch(r"v?\d+(\.\d+){0,3}[-\w.]*",meta["version"],re.I): raise AIResponseError("版本格式不合法")
    if not isinstance(meta["categories"],list): raise AIResponseError("分类必须为数组")
    for category in meta["categories"]:
        if not isinstance(category,dict) or set(category)!={"name","confidence","reason","evidence_refs"} or category["name"] not in CATEGORIES: raise AIResponseError("分类对象不合法")
        if not isinstance(category["confidence"],(int,float)) or not 0<=category["confidence"]<=1: raise AIResponseError("分类置信度不合法")
        _refs(category["evidence_refs"],evidence_ids)
    for key, allowed in STRUCTURED.items():
        if not isinstance(value[key],list): raise AIResponseError(f"{key} 必须是数组")
        for item in value[key]:
            if not isinstance(item,dict) or set(item)!=allowed: raise AIResponseError(f"{key} 元素字段不合法")
            if not isinstance(item["confidence"],(int,float)) or not 0<=item["confidence"]<=1: raise AIResponseError("专业记录置信度不合法")
            _refs(item["evidence_refs"],evidence_ids)
    return value

@dataclass
class AIAnalyzer:
    settings: Settings
    transport: Callable[[dict],dict]|None=None
    def analyze(self, context: str, evidence_items: list[dict]|None=None) -> dict|None:
        if not self.settings.ai_enabled or not self.settings.api_base_url or not self.settings.text_model:return None
        ids={x["evidence_id"] for x in evidence_items or []}; last=None
        for attempt in range(self.settings.retries+1):
            try:
                payload=self._payload(context,evidence_items or [])
                response=self.transport(payload) if self.transport else self._request(payload)
                return validate_ai_result(json.loads(response["choices"][0]["message"]["content"]),ids)
            except (KeyError,TypeError,json.JSONDecodeError,urllib.error.URLError,TimeoutError,AIResponseError) as exc:
                last=exc
                if attempt<self.settings.retries:time.sleep(min(2**attempt,3))
        raise AIResponseError(f"AI 分析失败：{type(last).__name__}")
    def _payload(self, context, evidence_items):
        data={"model":self.settings.text_model,"messages":[{"role":"system","content":"仅基于 evidence 生成 JSON。所有关键字段必须引用 evidence_refs；无证据填待人工确认。"},{"role":"user","content":json.dumps({"context":context,"evidence":evidence_items},ensure_ascii=False)}],"temperature":0}
        if self.settings.ai_schema_mode=="strict":data["response_format"]={"type":"json_schema","json_schema":AI_SCHEMA}
        return data
    def _request(self,payload):
        headers={"Content-Type":"application/json"}
        if self.settings.api_key:headers["Authorization"]=f"Bearer {self.settings.api_key}"
        request=urllib.request.Request(self.settings.api_base_url.rstrip("/")+"/chat/completions",data=json.dumps(payload).encode(),headers=headers,method="POST")
        with urllib.request.urlopen(request,timeout=self.settings.timeout_seconds) as response:return json.loads(response.read().decode())
    def embeddings(self,texts):
        if not self.settings.semantic_similarity_enabled or not self.settings.api_base_url or not self.settings.embedding_model:return None
        all_vectors=[]
        for start in range(0,len(texts),self.settings.embedding_batch_size):
            headers={"Content-Type":"application/json"};
            if self.settings.api_key:headers["Authorization"]=f"Bearer {self.settings.api_key}"
            request=urllib.request.Request(self.settings.api_base_url.rstrip("/")+"/embeddings",data=json.dumps({"model":self.settings.embedding_model,"input":texts[start:start+self.settings.embedding_batch_size]}).encode(),headers=headers,method="POST")
            with urllib.request.urlopen(request,timeout=self.settings.timeout_seconds) as response:all_vectors.extend(x["embedding"] for x in sorted(json.loads(response.read().decode())["data"],key=lambda x:x["index"]))
        return all_vectors

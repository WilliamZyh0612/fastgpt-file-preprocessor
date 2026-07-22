from __future__ import annotations
import json, re, time, urllib.error, urllib.request
from datetime import datetime
from dataclasses import dataclass
from dataclasses import field
from threading import Lock
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
_text_schema = {"type":"string", "maxLength":4000}
_short_text_schema = {"type":"string", "maxLength":128}
_refs_schema = {"type":"array", "minItems":1, "maxItems":50, "items":{"type":"string", "pattern":"^[a-f0-9]{16}$"}}
_confidence = {"type":"number", "minimum":0, "maximum":1}
_field_schema = {"type":"object", "additionalProperties":False, "required":["value", "confidence", "evidence_refs"], "properties":{"value":_text_schema, "confidence":_confidence, "evidence_refs":_refs_schema}}
_date_field_schema = {"type":"object", "additionalProperties":False, "required":["value", "confidence", "evidence_refs"], "properties":{"value":{"type":"string", "maxLength":32, "pattern":"^(待人工确认|\\d{4}-\\d{2}-\\d{2})$"}, "confidence":_confidence, "evidence_refs":_refs_schema}}
_version_field_schema = {"type":"object", "additionalProperties":False, "required":["value", "confidence", "evidence_refs"], "properties":{"value":{"type":"string", "maxLength":128, "pattern":"^(待人工确认|v?\\d+(?:\\.\\d+){0,3}[-A-Za-z0-9._]*)$"}, "confidence":_confidence, "evidence_refs":_refs_schema}}
_keyword_field_schema = {"type":"object", "additionalProperties":False, "required":["value", "confidence", "evidence_refs"], "properties":{"value":_short_text_schema, "confidence":_confidence, "evidence_refs":_refs_schema}}
_category = {"type":"object", "additionalProperties":False, "required":["name", "confidence", "reason", "evidence_refs"], "properties":{"name":{"type":"string", "enum":sorted(CATEGORIES)}, "confidence":_confidence, "reason":_text_schema, "evidence_refs":_refs_schema}}
_cnc_system = {"type":"object", "additionalProperties":False, "required":["brand", "model", "version", "confidence", "evidence_refs"], "properties":{"brand":_short_text_schema, "model":_short_text_schema, "version":_short_text_schema, "confidence":_confidence, "evidence_refs":_refs_schema}}
_metadata = {"type":"object", "additionalProperties":False, "required":sorted(DOC_FIELDS), "properties":{
    "categories":{"type":"array", "maxItems":20, "items":_category},
    "machine_models":{"type":"array", "maxItems":50, "items":_field_schema},
    "cnc_systems":{"type":"array", "maxItems":50, "items":_cnc_system},
    "visibility":_field_schema, "version":_version_field_schema, "release_date":_date_field_schema,
    "summary":_field_schema, "knowledge_points":{"type":"array", "maxItems":30, "items":_field_schema},
    "keywords":{"type":"array", "maxItems":30, "items":_keyword_field_schema},
}}
_record_schemas = {key:{"type":"array", "maxItems":100, "items":{"type":"object", "additionalProperties":False, "required":sorted(fields), "properties":{field:(_refs_schema if field=="evidence_refs" else _confidence if field=="confidence" else {"type":"boolean"} if field=="requires_engineer" else _text_schema) for field in fields}}} for key,fields in STRUCTURED.items()}
AI_SCHEMA = {"name":"knowledge_preprocessor_result","strict":True,"schema":{"type":"object","additionalProperties":False,"required":["document_metadata",*STRUCTURED],"properties":{"document_metadata":_metadata,**_record_schemas}}}

class AIResponseError(ValueError): pass

def _ensure_string(value: object, label: str, maximum: int = 4000) -> None:
    if not isinstance(value, str) or len(value)>maximum: raise AIResponseError(f"{label} 必须是长度受限字符串")
def _refs(value: object, ids: set[str] | None) -> None:
    if not isinstance(value,list) or not value or len(value)>50 or not all(isinstance(x,str) and re.fullmatch(r"[a-f0-9]{16}", x) for x in value): raise AIResponseError("evidence_refs 类型不合法")
    if ids is not None and not set(value).issubset(ids): raise AIResponseError("AI 返回不存在的 evidence_refs")

def _confidence_value(value: object, label: str) -> None:
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not 0 <= value <= 1:
        raise AIResponseError(f"{label} 置信度不合法")

def _field(value: object, label: str, ids: set[str] | None, *, date: bool = False, version: bool = False, short: bool = False) -> None:
    if not isinstance(value, dict) or set(value) != {"value", "confidence", "evidence_refs"}: raise AIResponseError(f"{label} 字段对象不合法")
    _ensure_string(value["value"], label, 128 if short else 4000); _confidence_value(value["confidence"], label); _refs(value["evidence_refs"], ids)
    if date:
        if value["value"] != "待人工确认":
            try: datetime.strptime(value["value"], "%Y-%m-%d")
            except ValueError as exc: raise AIResponseError("日期必须为有效 yyyy-mm-dd") from exc
    if version and value["value"] != "待人工确认" and not re.fullmatch(r"v?\d+(\.\d+){0,3}[-\w.]*", value["value"], re.I): raise AIResponseError("版本格式不合法")

def validate_ai_result(value: object, evidence_ids: set[str] | None = None) -> dict:
    if not isinstance(value,dict) or set(value)!={"document_metadata",*STRUCTURED}: raise AIResponseError("AI 顶层字段不符合严格 Schema")
    meta=value["document_metadata"]
    if not isinstance(meta,dict) or set(meta)!=DOC_FIELDS: raise AIResponseError("document_metadata 字段不合法")
    for name in ("visibility", "summary"): _field(meta[name], name, evidence_ids)
    _field(meta["version"], "version", evidence_ids, version=True); _field(meta["release_date"], "release_date", evidence_ids, date=True)
    if not isinstance(meta["categories"],list): raise AIResponseError("分类必须为数组")
    for category in meta["categories"]:
        if not isinstance(category,dict) or set(category)!={"name","confidence","reason","evidence_refs"} or category["name"] not in CATEGORIES: raise AIResponseError("分类对象不合法")
        _ensure_string(category["reason"], "分类理由"); _confidence_value(category["confidence"], "分类")
        _refs(category["evidence_refs"],evidence_ids)
    if not isinstance(meta["machine_models"], list) or not isinstance(meta["knowledge_points"], list) or not isinstance(meta["keywords"], list) or not isinstance(meta["cnc_systems"], list): raise AIResponseError("元数据数组类型不合法")
    for item in meta["machine_models"]: _field(item, "machine_models", evidence_ids)
    for item in meta["knowledge_points"]: _field(item, "knowledge_points", evidence_ids)
    for item in meta["keywords"]: _field(item, "keywords", evidence_ids, short=True)
    for item in meta["cnc_systems"]:
        if not isinstance(item, dict) or set(item) != {"brand", "model", "version", "confidence", "evidence_refs"}: raise AIResponseError("数控系统对象不合法")
        for field in ("brand", "model", "version"): _ensure_string(item[field], f"数控系统{field}", 128)
        _confidence_value(item["confidence"], "数控系统"); _refs(item["evidence_refs"], evidence_ids)
    for key, allowed in STRUCTURED.items():
        if not isinstance(value[key],list): raise AIResponseError(f"{key} 必须是数组")
        for item in value[key]:
            if not isinstance(item,dict) or set(item)!=allowed: raise AIResponseError(f"{key} 元素字段不合法")
            for field, field_value in item.items():
                if field not in {"evidence_refs", "confidence", "requires_engineer"}: _ensure_string(field_value, field)
            if "requires_engineer" in item and not isinstance(item["requires_engineer"], bool): raise AIResponseError("requires_engineer 必须是布尔值")
            _confidence_value(item["confidence"], "专业记录")
            _refs(item["evidence_refs"],evidence_ids)
    return value

@dataclass
class AIAnalyzer:
    settings: Settings
    transport: Callable[[dict],dict]|None=None
    usage: dict[str, int] = field(default_factory=lambda: {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0})
    _usage_lock: Lock = field(default_factory=Lock, repr=False)
    def analyze(self, context: str, evidence_items: list[dict]|None=None) -> dict|None:
        if not self.settings.ai_enabled or not self.settings.api_base_url or not self.settings.text_model:return None
        ids={x["evidence_id"] for x in evidence_items or []}; last=None
        for attempt in range(self.settings.retries+1):
            try:
                payload=self._payload(context,evidence_items or [])
                response=self.transport(payload) if self.transport else self._request(payload)
                self._record_usage(response.get("usage", {}))
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
    def _record_usage(self, usage: object) -> None:
        if not isinstance(usage, dict): return
        with self._usage_lock:
            for key in self.usage:
                value = usage.get(key, 0)
                if isinstance(value, int) and value >= 0: self.usage[key] += value
    def embeddings(self,texts):
        if not self.settings.semantic_similarity_enabled or not self.settings.api_base_url or not self.settings.embedding_model:return None
        all_vectors=[]
        for start in range(0,len(texts),self.settings.embedding_batch_size):
            headers={"Content-Type":"application/json"};
            if self.settings.api_key:headers["Authorization"]=f"Bearer {self.settings.api_key}"
            request=urllib.request.Request(self.settings.api_base_url.rstrip("/")+"/embeddings",data=json.dumps({"model":self.settings.embedding_model,"input":texts[start:start+self.settings.embedding_batch_size]}).encode(),headers=headers,method="POST")
            with urllib.request.urlopen(request,timeout=self.settings.timeout_seconds) as response:all_vectors.extend(x["embedding"] for x in sorted(json.loads(response.read().decode())["data"],key=lambda x:x["index"]))
        return all_vectors

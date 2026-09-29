from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
import uuid
from .enums import DocumentStatus, JobState, OrderListStatus, SortMode, SeparatorMode, ErrorPolicy

@dataclass
class DocumentStats:
    paragraphs:int=0; tables:int=0; images:int=0; sections:int=0; headers:int=0; footers:int=0; footnotes:int=0; endnotes:int=0; comments:int=0; bookmarks:int=0; fields:int=0; embedded_objects:int=0

@dataclass
class DocumentItem:
    source_path:str
    id:str=field(default_factory=lambda:str(uuid.uuid4()))
    filename:str=""; extension:str=""; detected_format:str="UNKNOWN"; size_bytes:int=0; modified_time:float=0.0; source_folder:str=""
    enabled:bool=True; manual_order:int=0
    health_status:DocumentStatus=DocumentStatus.NEW; merge_status:DocumentStatus=DocumentStatus.NEW
    warnings:list[str]=field(default_factory=list); errors:list[str]=field(default_factory=list)
    binary_hash:str|None=None; content_hash:str|None=None; encoding:str|None=None; encoding_confidence:float|None=None; open_test_status:str|None=None
    stats:DocumentStats=field(default_factory=DocumentStats)

    def hydrate_from_path(self):
        p=Path(self.source_path); self.filename=p.name; self.extension=p.suffix.lower(); self.source_folder=str(p.parent)
        if p.exists():
            s=p.stat(); self.size_bytes=s.st_size; self.modified_time=s.st_mtime

    @property
    def eligible_for_merge(self):
        return self.enabled and self.health_status in {DocumentStatus.OK,DocumentStatus.WARNING,DocumentStatus.READY} and not self.errors

    def to_dict(self):
        d=asdict(self); d["health_status"]=self.health_status.value; d["merge_status"]=self.merge_status.value; return d

    @classmethod
    def from_dict(cls,data):
        data=dict(data); stats=data.pop("stats",{}) or {}; x=cls(**data); x.health_status=DocumentStatus(x.health_status); x.merge_status=DocumentStatus(x.merge_status); x.stats=DocumentStats(**stats); return x

@dataclass
class OrderEntry:
    """Viens CSV/JSON secības saraksta ieraksts (pirms matching)."""
    order:int=0
    requested:str=""
    raw:str=""
    line:int=0

    def to_dict(self): return {"order":self.order,"requested":self.requested}

    @classmethod
    def from_dict(cls,data):
        data=dict(data); return cls(order=int(data.get("order",0) or 0),requested=str(data.get("requested","")))


@dataclass
class OrderMatch:
    """Saraksta ieraksta matching rezultāts (GUI priekšskatījumam un reportam)."""
    order:int=0
    requested:str=""
    status:OrderListStatus=OrderListStatus.MISSING
    resolved_path:str=""
    resolved_name:str=""
    detected_format:str=""
    item_id:str=""
    rule:str=""
    message:str=""

    @property
    def ok(self): return self.status==OrderListStatus.FOUND

    def to_dict(self):
        return {"order":self.order,"requested":self.requested,"status":self.status.value,
                "resolved_path":self.resolved_path,"resolved_name":self.resolved_name,
                "detected_format":self.detected_format,"item_id":self.item_id,"rule":self.rule,
                "message":self.message}

    @classmethod
    def from_dict(cls,data):
        data=dict(data)
        return cls(order=int(data.get("order",0) or 0),requested=str(data.get("requested","")),
                   status=OrderListStatus(data.get("status",OrderListStatus.MISSING.value)),
                   resolved_path=str(data.get("resolved_path","")),resolved_name=str(data.get("resolved_name","")),
                   detected_format=str(data.get("detected_format","")),item_id=str(data.get("item_id","")),
                   rule=str(data.get("rule","")),message=str(data.get("message","")))


@dataclass
class MergeOptions:
    sort_mode:SortMode=SortMode.NATURAL_NAME
    separator:SeparatorMode=SeparatorMode.PAGE_BREAK
    preserve_source_sections:bool=True
    insert_control_markers:bool=True
    detect_duplicates:bool=True
    error_policy:ErrorPolicy=ErrorPolicy.RETRY_ONCE_THEN_SKIP
    word_timeout_seconds:int=30
    # Faila nosaukums kā trekns virsraksts pirms katra dokumenta (kā pārbaudītajā skriptā).
    insert_titles:bool=True
    # Protected View: noņem Zone.Identifier izvadei un izmanto atbloķētu avota kopiju.
    unlock_protected_view:bool=True
    # CSV/JSON secības saraksts (absolūtā merge secība, natural sort netiek pielietots).
    order_list_enabled:bool=False
    order_list_path:str=""
    order_list_entries:list[str]=field(default_factory=list)
    strict_order_mode:bool=False
    # Atvērt izvades DOCX pēc veiksmīgas apvienošanas (noklusēti IZSLĒGTS).
    open_output_when_finished:bool=False

@dataclass
class Project:
    name:str="Jauns projekts"; id:str=field(default_factory=lambda:str(uuid.uuid4())); version:int=1
    created_at:str=field(default_factory=lambda:datetime.now().isoformat(timespec="seconds")); updated_at:str=field(default_factory=lambda:datetime.now().isoformat(timespec="seconds"))
    output_path:str=""; items:list[DocumentItem]=field(default_factory=list); options:MergeOptions=field(default_factory=MergeOptions); state:JobState=JobState.IDLE
    def to_dict(self):
        return {"name":self.name,"id":self.id,"version":self.version,"created_at":self.created_at,"updated_at":datetime.now().isoformat(timespec="seconds"),"output_path":self.output_path,"state":self.state.value,"options":self.options_to_dict(),"items":[x.to_dict() for x in self.items]}

    def options_to_dict(self):
        o=self.options
        return {"sort_mode":o.sort_mode.value,"separator":o.separator.value,"preserve_source_sections":o.preserve_source_sections,"insert_control_markers":o.insert_control_markers,"detect_duplicates":o.detect_duplicates,"error_policy":o.error_policy.value,"word_timeout_seconds":o.word_timeout_seconds,"insert_titles":o.insert_titles,"unlock_protected_view":o.unlock_protected_view,"order_list_enabled":o.order_list_enabled,"order_list_path":o.order_list_path,"order_list_entries":list(o.order_list_entries),"strict_order_mode":o.strict_order_mode,"open_output_when_finished":o.open_output_when_finished}
    @classmethod
    def from_dict(cls,data):
        o=data.get("options",{}); p=cls(name=data.get("name","Projekts"),id=data.get("id",str(uuid.uuid4())),version=data.get("version",1),created_at=data.get("created_at",datetime.now().isoformat(timespec="seconds")),updated_at=data.get("updated_at",datetime.now().isoformat(timespec="seconds")),output_path=data.get("output_path",""),state=JobState(data.get("state",JobState.IDLE.value)),options=MergeOptions(sort_mode=SortMode(o.get("sort_mode",SortMode.NATURAL_NAME.value)),separator=SeparatorMode(o.get("separator",SeparatorMode.PAGE_BREAK.value)),preserve_source_sections=o.get("preserve_source_sections",True),insert_control_markers=o.get("insert_control_markers",True),detect_duplicates=o.get("detect_duplicates",True),error_policy=ErrorPolicy(o.get("error_policy",ErrorPolicy.RETRY_ONCE_THEN_SKIP.value)),word_timeout_seconds=int(o.get("word_timeout_seconds",30)),insert_titles=bool(o.get("insert_titles",True)),unlock_protected_view=bool(o.get("unlock_protected_view",True))))
        p.items=[DocumentItem.from_dict(x) for x in data.get("items",[])]
        p.options.order_list_enabled=bool(o.get("order_list_enabled",False))
        p.options.order_list_path=str(o.get("order_list_path","") or "")
        p.options.order_list_entries=[str(x) for x in (o.get("order_list_entries") or [])]
        p.options.strict_order_mode=bool(o.get("strict_order_mode",False))
        p.options.open_output_when_finished=bool(o.get("open_output_when_finished",False))
        return p

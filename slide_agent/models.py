from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, Field, StringConstraints, model_validator

Id = Annotated[str, StringConstraints(pattern=r"^[a-zA-Z][a-zA-Z0-9_-]{0,39}$")]
Short = Annotated[str, StringConstraints(min_length=1, max_length=500)]


class Model(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True, allow_inf_nan=False)


class Segment(Model):
    id: Id
    text: Annotated[str, Field(min_length=1, max_length=50000)]
    citation: Annotated[str, Field(max_length=2000)] = ""
    group: Id = "document"


class Asset(Model):
    id: Id
    path: Annotated[str, Field(min_length=1, max_length=300)]
    caption: Annotated[str, Field(max_length=500)] = ""
    sha256: Annotated[str, Field(pattern=r"^[a-f0-9]{64}$")]


class Source(Model):
    version: Literal["1"] = "1"
    title: Short
    segments: Annotated[list[Segment], Field(min_length=1, max_length=500)]
    images: Annotated[list[Asset], Field(max_length=100)] = []

    @model_validator(mode="after")
    def unique_ids(self):
        for collection in (self.segments, self.images):
            ids = [entry.id for entry in collection]
            if len(ids) != len(set(ids)):
                raise ValueError("source and image IDs must be unique within each collection")
        return self


class Ref(Model):
    source_id: Id
    quote: Annotated[str, Field(min_length=1, max_length=50000)]


class Text(Model):
    text: Annotated[str, Field(min_length=1, max_length=20000)]
    refs: Annotated[list[Ref], Field(min_length=1, max_length=30)]
    mode: Literal["verbatim", "paraphrase"] = "verbatim"


class Items(Model):
    items: Annotated[list[Text], Field(max_length=100)]


class ImageContents(Items):
    image_id: Id
    image_mode: Literal["fit", "crop"] = "fit"


class Comparison(Model):
    left_title: Text
    right_title: Text
    left: Annotated[list[Text], Field(min_length=1, max_length=30)]
    right: Annotated[list[Text], Field(min_length=1, max_length=30)]


class Process(Model):
    steps: Annotated[list[Text], Field(min_length=2, max_length=20)]


class Table(Model):
    headers: Annotated[list[Text], Field(min_length=2, max_length=3)]
    rows: Annotated[list[list[Text]], Field(min_length=1, max_length=100)]

    @model_validator(mode="after")
    def rectangular(self):
        if any(len(row) != len(self.headers) for row in self.rows):
            raise ValueError("every table row must match the header count")
        return self


class Datum(Model):
    value: Annotated[float, Field(ge=0, le=1e12)]
    refs: Annotated[list[Ref], Field(min_length=1, max_length=10)]


class Series(Model):
    name: Text
    values: Annotated[list[Datum], Field(min_length=1, max_length=6)]


class Chart(Model):
    categories: Annotated[list[Text], Field(min_length=1, max_length=6)]
    series: Annotated[list[Series], Field(min_length=1, max_length=2)]
    note: Text

    @model_validator(mode="after")
    def dimensions(self):
        if any(len(s.values) != len(self.categories) for s in self.series):
            raise ValueError("chart values must match category count")
        return self


class SlideBase(Model):
    id: Id
    origin: Id
    title: Text
    lead: Text | None = None
    rationale: Annotated[str, Field(min_length=1, max_length=2000)]


class TitleSlide(SlideBase):
    layout_id: Literal["title"]
    contents: Items


class BulletSlide(SlideBase):
    layout_id: Literal["bullets"]
    contents: Items


class ImageSlide(SlideBase):
    layout_id: Literal["text_image"]
    contents: ImageContents


class ComparisonSlide(SlideBase):
    layout_id: Literal["comparison"]
    contents: Comparison


class ProcessSlide(SlideBase):
    layout_id: Literal["process"]
    contents: Process


class TableSlide(SlideBase):
    layout_id: Literal["table"]
    contents: Table


class ChartSlide(SlideBase):
    layout_id: Literal["chart"]
    contents: Chart


class ClosingSlide(SlideBase):
    layout_id: Literal["closing"]
    contents: Items


class SignedDatum(Datum):
    value: Annotated[float, Field(ge=-1e12, le=1e12)]


class TemplateSeries(Model):
    name: Text
    values: Annotated[list[SignedDatum], Field(min_length=1, max_length=60)]
    x_values: Annotated[list[SignedDatum], Field(max_length=60)] = []


class ElapsedYears(Datum):
    value: Annotated[float, Field(gt=0, le=1000)]


class TemplateChart(Model):
    categories: Annotated[list[Text], Field(max_length=60)] = []
    series: Annotated[list[TemplateSeries], Field(min_length=1, max_length=4)]
    elapsed_years: ElapsedYears | None = None


class CatalogContents(Model):
    texts: dict[Id, Text]
    charts: dict[Id, TemplateChart] = {}
    metrics: dict[Id, SignedDatum] = {}
    states: dict[Id, Text] = {}


from .catalog import REFERENCE_IDS, EDITORIAL_IDS, RELATION_IDS


class CatalogSlide(SlideBase):
    layout_id: Literal[REFERENCE_IDS]
    variant: Literal["warm", "cool"] = "warm"
    font_profile: Literal["source", "meiryo", "noto", "hiragino"] = "source"
    contents: CatalogContents

    @model_validator(mode="after")
    def no_lead(self):
        if self.lead is not None:
            raise ValueError("catalog templates have fixed slots; lead is not supported")
        return self


class TemplateImage(Model):
    image_id: Id
    image_mode: Literal['fit', 'crop'] = 'fit'


class EditorialContents(CatalogContents):
    images: dict[Id, TemplateImage] = {}


class EditorialEmphasis(Model):
    item_id: Literal['item_1']
    refs: Annotated[list[Ref], Field(min_length=1, max_length=10)]
    reason: Annotated[str, Field(min_length=1, max_length=1000)]


class EditorialSlide(CatalogSlide):
    layout_id: Literal[EDITORIAL_IDS]
    contents: EditorialContents
    emphasis: EditorialEmphasis | None = None
    allowed_layouts: Annotated[list[Id], Field(max_length=100)] = []
    constraints_reason: Annotated[str, Field(min_length=1, max_length=1000)] | None = None
    repetition_reason: Annotated[str, Field(min_length=1, max_length=1000)] | None = None
    row_alignments: dict[Id, Literal['top', 'middle']] = {}

    @model_validator(mode='after')
    def explain_constraints(self):
        if self.allowed_layouts and not self.constraints_reason:
            raise ValueError('allowed_layouts requires an explicit content/reading-order constraints_reason')
        return self


class RelationNode(Model):
    id: Id
    label: Text


class RelationEdge(Model):
    id: Id
    source: Id
    target: Id
    label: Text


class RelationBoundary(Model):
    id: Id
    kind: Literal['contains', 'same_owner']
    label: Text
    members: Annotated[list[Id], Field(min_length=1, max_length=5)]


class Network(Model):
    nodes: Annotated[list[RelationNode], Field(min_length=2, max_length=5)]
    edges: Annotated[list[RelationEdge], Field(min_length=1, max_length=8)]
    boundaries: Annotated[list[RelationBoundary], Field(max_length=1)] = []


class ComparisonMethod(Model):
    id: Id
    label: Text
    values: Annotated[list[Text], Field(min_length=3, max_length=3)]


class ThreeMethods(Model):
    axes: Annotated[list[Text], Field(min_length=3, max_length=3)]
    methods: Annotated[list[ComparisonMethod], Field(min_length=3, max_length=3)]


class RelationContents(EditorialContents):
    network: Network | None = None
    comparison: ThreeMethods | None = None


class RelationSlide(EditorialSlide):
    layout_id: Literal[RELATION_IDS]
    contents: RelationContents


Slide = Annotated[
    TitleSlide | BulletSlide | ImageSlide | ComparisonSlide | ProcessSlide | TableSlide | ChartSlide | ClosingSlide | CatalogSlide | EditorialSlide | RelationSlide,
    Field(discriminator="layout_id"),
]


class Omission(Model):
    source_id: Id
    reason: Annotated[str, Field(min_length=1, max_length=500)]


class DisplayCitation(Model):
    source_ids: Annotated[list[Id], Field(min_length=1, max_length=500)]
    title: Annotated[str, Field(min_length=1, max_length=100)]
    url: Annotated[str, StringConstraints(pattern=r'^https?://[^\s]+$', max_length=2000)]
    version: Annotated[str, Field(max_length=60)] = ''
    accessed_on: Annotated[str, StringConstraints(pattern=r'^\d{4}-\d{2}-\d{2}$')] | None = None


class DeckDisplay(Model):
    mode: Literal['reader', 'qa'] = 'reader'
    footer: Annotated[str, Field(max_length=100)] = ''
    citations: Annotated[list[DisplayCitation], Field(max_length=50)] = []


class Plan(Model):
    version: Literal["1"] = "1"
    source_sha256: Annotated[str, Field(pattern=r"^[a-f0-9]{64}$")]
    title: Short
    slides: Annotated[list[Slide], Field(min_length=1, max_length=100)]
    omissions: Annotated[list[Omission], Field(max_length=500)] = []
    display: DeckDisplay = Field(default_factory=DeckDisplay)

    @model_validator(mode="after")
    def unique_ids(self):
        if len({s.id for s in self.slides}) != len(self.slides):
            raise ValueError("slide IDs must be unique")
        if len({s.source_id for s in self.omissions}) != len(self.omissions):
            raise ValueError("omission IDs must be unique")
        return self


Color = Annotated[str, Field(pattern=r"^[A-Fa-f0-9]{6}$")]


class Theme(Model):
    font_family: Annotated[str, Field(min_length=1, max_length=80)] = "Meiryo"
    background: Color = "F5F7FA"
    foreground: Color = "13233B"
    accent: Color = "007C83"
    secondary: Color = "47699B"
    muted: Color = "5B687A"
    surface: Color = "FFFFFF"
    surface_alt: Color = "E8EEF4"
    title_pt: Annotated[int, Field(ge=28, le=36)] = 32
    body_pt: Annotated[int, Field(ge=22, le=28)] = 24
    table_pt: Annotated[int, Field(ge=18, le=22)] = 20
    footnote_pt: Annotated[int, Field(ge=11, le=14)] = 12
    lead_mode: bool = False

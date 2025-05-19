from collections.abc import Iterator
from enum import StrEnum

from pydantic import BaseModel, RootModel


class Weather(BaseModel):
    cloud: int
    rain: int
    wind: int
    visibility: int


CLOUD_KEY = ["", "0-33%", "33-66%", "66-100%"]
RAIN_KEY = ["", "none", "drizzle", "showers"]
WIND_KEY = ["", "calm", "light", "breezy"]
VISIBILITY_KEY = ["", "good", "moderate", "poor"]


class BtoSpeciesCode(StrEnum):  # noqa: F821
    """
    Enumeration of BTO Species Codes from the provided image.

    Each member's name is a Python-friendly representation of the species name,
    and its value is the corresponding BTO code string.
    """

    BARN_OWL = "BO"
    BLACKBIRD = "B."
    BLACKCAP = "BC"
    BLACK_HEADED_GULL = "BH"
    BLUE_TIT = "BT"
    BULLFINCH = "BF"
    BUZZARD = "BZ"
    CANADA_GOOSE = "CG"
    CARRION_CROW = "C."
    CHAFFINCH = "CH"
    CHIFFCHAFF = "CC"
    COAL_TIT = "CT"
    COLLARED_DOVE = "CD"
    COMMON_GULL = "CM"
    CORMORANT = "CA"
    CROSSBILL_COMMON = "CR"
    CUCKOO = "CK"
    CURLEW = "CU"
    DUNLIN = "DN"
    DUNNOCK = "D."
    FERAL_PIGEON = "FP"
    FERAL_HYBRID_GOOSE = "ZL"
    FERAL_HYBRID_MALLARD_TYPE = "ZF"
    GARDEN_WARBLER = "GW"
    GOLDCREST = "GC"
    GOLDEN_PLOVER = "GP"
    GOLDFINCH = "GO"
    GREAT_SPOTTED_WOODPECKER = "GS"
    GRASSHOPPER_WARBLER = "GH"
    GREAT_TIT = "GT"
    GREEN_WOODPECKER = "G."
    GREENFINCH = "GR"
    GREY_WAGTAIL = "GL"
    GREYLAG_GOOSE = "GJ"
    GREY_HERON = "H."
    HEN_HARRIER = "HH"
    HERRING_GULL = "HG"
    HOBBY = "HY"
    HOUSE_MARTIN = "HM"
    HOUSE_SPARROW = "HS"
    JAY = "J."
    JACKDAW = "JD"
    KESTREL = "K."
    LAPWING = "L."
    LESSER_BLACK_BACKED_GULL = "LB"
    LINNET = "LI"
    LITTLE_OWL = "LO"
    LONG_EARED_OWL = "LE"
    LONG_TAILED_TIT = "LT"
    MAGPIE = "MG"
    MALLARD = "MA"
    MARSH_HARRIER = "MR"
    MEADOW_PIPIT = "MP"
    MERLIN = "ML"
    MISTLE_THRUSH = "M."
    MOORHEN = "MH"
    NIGHTJAR = "NJ"
    NUTHATCH = "NH"
    OYSTERCATCHER = "OC"
    PEREGRINE = "PE"
    PHEASANT = "PH"
    PIED_FLYCATCHER = "PF"
    PIED_WAGTAIL = "PW"
    RAVEN = "RN"
    RED_GROUSE = "RG"
    RED_KITE = "KT"
    RED_LEGGED_PARTRIDGE = "RL"
    REDPOLL_LESSER = "LR"
    REDSHANK = "RK"
    REDSTART = "RT"
    RING_OUZEL = "RZ"
    ROBIN = "R."
    ROOK = "RO"
    SAND_MARTIN = "SM"
    SHELDUCK = "SU"
    SHORT_EARED_OWL = "SE"
    SISKIN = "SK"
    SKYLARK = "S."
    SNIPE = "SN"
    SONG_THRUSH = "ST"
    SPARROWHAWK = "SH"
    STOCK_DOVE = "SD"
    STONECHAT = "SC"
    SPOTTED_FLYCATCHER = "SF"
    STARLING = "SG"
    SWALLOW = "SL"
    SWIFT = "SI"
    TAWNY_OWL = "TO"
    TREE_PIPIT = "TP"
    TREECREEPER = "TC"
    WHEATEAR = "W."
    WHINCHAT = "WC"
    WILLOW_WARBLER = "WW"
    WOODCOCK = "WK"
    WOODPIGEON = "WP"
    WREN = "WR"
    YELLOW_WAGTAIL = "YW"
    YELLOWHAMMER = "Y."

    @property
    def bird_name(self):
        return f"{' '.join(word.capitalize() for word in self.name.lower().split('_'))}"


class Sighting(BaseModel):
    count: int | None
    code: BtoSpeciesCode
    comment: str | None  # to be ignored


class Segment(BaseModel):
    number: int
    left: list[Sighting]
    right: list[Sighting]
    start_coordinate: str | None
    end_coordinate: str | None


class SurveyData(BaseModel):
    observer_name: str
    transect_number: int
    visit_date: str
    weather_code: Weather
    first_segment_start_time: str
    first_segment_end_time: str
    second_segment_start_time: str
    second_segment_end_time: str

    segments: list[Segment]


class Surveys(RootModel[list[SurveyData]]):
    def __iter__(self) -> Iterator[SurveyData]:  # type:ignore[override]
        return iter(self.root)

    def append(self, survey_data: SurveyData) -> None:
        self.root.append(survey_data)


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Set,
    String,
    Bool,
    MinMax,
    Integer
)
from openpyxl.descriptors.excel import (
    HexBinary,
    TextPoint,
    Coordinate
)
from .shapes import (
    LineProperties,
    Color,
    Scene3D
)

class OfficeArtExtensionList:
    pass


class EmbeddedWAVAudioFile(Serialisable):

    name = Typed(expected_type=String, allow_none=True)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Hyperlink(Serialisable):

    invalidUrl = Typed(expected_type=String, allow_none=True)
    action = Typed(expected_type=String, allow_none=True)
    tgtFrame = Typed(expected_type=String, allow_none=True)
    tooltip = Typed(expected_type=String, allow_none=True)
    history = Typed(expected_type=Bool, allow_none=True)
    highlightClick = Typed(expected_type=Bool, allow_none=True)
    endSnd = Typed(expected_type=Bool, allow_none=True)
    snd = Typed(expected_type=EmbeddedWAVAudioFile, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 invalidUrl=None,
                 action=None,
                 tgtFrame=None,
                 tooltip=None,
                 history=None,
                 highlightClick=None,
                 endSnd=None,
                 snd=None,
                 extLst=None,
                ):
        self.invalidUrl = invalidUrl
        self.action = action
        self.tgtFrame = tgtFrame
        self.tooltip = tooltip
        self.history = history
        self.highlightClick = highlightClick
        self.endSnd = endSnd
        self.snd = snd
        self.extLst = extLst


class TextFont(Serialisable):

    typeface = Typed(expected_type=String())
    panose = Typed(expected_type=HexBinary, allow_none=True)
    pitchFamily = Typed(expected_type=MinMax, allow_none=True)
    charset = Typed(expected_type=MinMax, allow_none=True)

    def __init__(self,
                 typeface=None,
                 panose=None,
                 pitchFamily=None,
                 charset=None,
                ):
        self.typeface = typeface
        self.panose = panose
        self.pitchFamily = pitchFamily
        self.charset = charset


class TextCharacterProperties(Serialisable):

    kumimoji = Typed(expected_type=Bool, allow_none=True)
    lang = Typed(expected_type=String, allow_none=True)
    altLang = Typed(expected_type=String, allow_none=True)
    sz = Typed(expected_type=Integer())
    b = Typed(expected_type=Bool, allow_none=True)
    i = Typed(expected_type=Bool, allow_none=True)
    u = Typed(expected_type=Set(values=(['none', 'words', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl'])))
    strike = Typed(expected_type=Set(values=(['noStrike', 'sngStrike', 'dblStrike'])))
    kern = Typed(expected_type=Integer())
    cap = Typed(expected_type=Set(values=(['none', 'small', 'all'])))
    spc = Typed(expected_type=TextPoint, allow_none=True)
    normalizeH = Typed(expected_type=Bool, allow_none=True)
    baseline = Typed(expected_type=String, allow_none=True)
    noProof = Typed(expected_type=Bool, allow_none=True)
    dirty = Typed(expected_type=Bool, allow_none=True)
    err = Typed(expected_type=Bool, allow_none=True)
    smtClean = Typed(expected_type=Bool, allow_none=True)
    smtId = Typed(expected_type=Integer, allow_none=True)
    bmk = Typed(expected_type=String, allow_none=True)
    ln = Typed(expected_type=LineProperties, allow_none=True)
    highlight = Typed(expected_type=Color, allow_none=True)
    latin = Typed(expected_type=TextFont, allow_none=True)
    ea = Typed(expected_type=TextFont, allow_none=True)
    cs = Typed(expected_type=TextFont, allow_none=True)
    sym = Typed(expected_type=TextFont, allow_none=True)
    hlinkClick = Typed(expected_type=Hyperlink, allow_none=True)
    hlinkMouseOver = Typed(expected_type=Hyperlink, allow_none=True)
    rtl = Typed(expected_type=Bool, allow_none=True, nested=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 kumimoji=None,
                 lang=None,
                 altLang=None,
                 sz=None,
                 b=None,
                 i=None,
                 u=None,
                 strike=None,
                 kern=None,
                 cap=None,
                 spc=None,
                 normalizeH=None,
                 baseline=None,
                 noProof=None,
                 dirty=None,
                 err=None,
                 smtClean=None,
                 smtId=None,
                 bmk=None,
                 ln=None,
                 highlight=None,
                 latin=None,
                 ea=None,
                 cs=None,
                 sym=None,
                 hlinkClick=None,
                 hlinkMouseOver=None,
                 rtl=None,
                 extLst=None,
                ):
        self.kumimoji = kumimoji
        self.lang = lang
        self.altLang = altLang
        self.sz = sz
        self.b = b
        self.i = i
        self.u = u
        self.strike = strike
        self.kern = kern
        self.cap = cap
        self.spc = spc
        self.normalizeH = normalizeH
        self.baseline = baseline
        self.noProof = noProof
        self.dirty = dirty
        self.err = err
        self.smtClean = smtClean
        self.smtId = smtId
        self.bmk = bmk
        self.ln = ln
        self.highlight = highlight
        self.latin = latin
        self.ea = ea
        self.cs = cs
        self.sym = sym
        self.hlinkClick = hlinkClick
        self.hlinkMouseOver = hlinkMouseOver
        self.rtl = rtl
        self.extLst = extLst


class TextTabStop(Serialisable):

    pos = Typed(expected_type=Coordinate, allow_none=True)
    algn = Typed(expected_type=Set(values=(['l', 'ctr', 'r', 'dec'])))

    def __init__(self,
                 pos=None,
                 algn=None,
                ):
        self.pos = pos
        self.algn = algn


class TextTabStopList(Serialisable):

    tab = Typed(expected_type=TextTabStop, allow_none=True)

    def __init__(self,
                 tab=None,
                ):
        self.tab = tab


class TextSpacing(Serialisable):

    pass

class TextParagraphProperties(Serialisable):

    marL = Typed(expected_type=Coordinate)
    marR = Typed(expected_type=Coordinate)
    lvl = Typed(expected_type=Integer())
    indent = Typed(expected_type=Coordinate)
    algn = Typed(expected_type=Set(values=(['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist'])))
    defTabSz = Typed(expected_type=Coordinate, allow_none=True)
    rtl = Typed(expected_type=Bool, allow_none=True)
    eaLnBrk = Typed(expected_type=Bool, allow_none=True)
    fontAlgn = Typed(expected_type=Set(values=(['auto', 't', 'ctr', 'base', 'b'])))
    latinLnBrk = Typed(expected_type=Bool, allow_none=True)
    hangingPunct = Typed(expected_type=Bool, allow_none=True)
    lnSpc = Typed(expected_type=TextSpacing, allow_none=True)
    spcBef = Typed(expected_type=TextSpacing, allow_none=True)
    spcAft = Typed(expected_type=TextSpacing, allow_none=True)
    tabLst = Typed(expected_type=TextTabStopList, allow_none=True)
    defRPr = Typed(expected_type=TextCharacterProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 marL=None,
                 marR=None,
                 lvl=None,
                 indent=None,
                 algn=None,
                 defTabSz=None,
                 rtl=None,
                 eaLnBrk=None,
                 fontAlgn=None,
                 latinLnBrk=None,
                 hangingPunct=None,
                 lnSpc=None,
                 spcBef=None,
                 spcAft=None,
                 tabLst=None,
                 defRPr=None,
                 extLst=None,
                ):
        self.marL = marL
        self.marR = marR
        self.lvl = lvl
        self.indent = indent
        self.algn = algn
        self.defTabSz = defTabSz
        self.rtl = rtl
        self.eaLnBrk = eaLnBrk
        self.fontAlgn = fontAlgn
        self.latinLnBrk = latinLnBrk
        self.hangingPunct = hangingPunct
        self.lnSpc = lnSpc
        self.spcBef = spcBef
        self.spcAft = spcAft
        self.tabLst = tabLst
        self.defRPr = defRPr
        self.extLst = extLst


class TextListStyle(Serialisable):

    defPPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl1pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl2pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl3pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl4pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl5pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl6pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl7pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl8pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    lvl9pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 defPPr=None,
                 lvl1pPr=None,
                 lvl2pPr=None,
                 lvl3pPr=None,
                 lvl4pPr=None,
                 lvl5pPr=None,
                 lvl6pPr=None,
                 lvl7pPr=None,
                 lvl8pPr=None,
                 lvl9pPr=None,
                 extLst=None,
                ):
        self.defPPr = defPPr
        self.lvl1pPr = lvl1pPr
        self.lvl2pPr = lvl2pPr
        self.lvl3pPr = lvl3pPr
        self.lvl4pPr = lvl4pPr
        self.lvl5pPr = lvl5pPr
        self.lvl6pPr = lvl6pPr
        self.lvl7pPr = lvl7pPr
        self.lvl8pPr = lvl8pPr
        self.lvl9pPr = lvl9pPr
        self.extLst = extLst


class TextParagraph(Serialisable):

    pPr = Typed(expected_type=TextParagraphProperties, allow_none=True)
    endParaRPr = Typed(expected_type=TextCharacterProperties, allow_none=True)

    def __init__(self,
                 pPr=None,
                 endParaRPr=None,
                ):
        self.pPr = pPr
        self.endParaRPr = endParaRPr


class GeomGuide(Serialisable):

    name = Typed(expected_type=String())
    fmla = Typed(expected_type=String())

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    gd = Typed(expected_type=GeomGuide, allow_none=True)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class PresetTextShape(Serialisable):

    prst = Typed(expected_type=Set(values=(['textNoShape', 'textPlain',
                                            'textStop', 'textTriangle', 'textTriangleInverted', 'textChevron',
                                            'textChevronInverted', 'textRingInside', 'textRingOutside', 'textArchUp',
                                            'textArchDown', 'textCircle', 'textButton', 'textArchUpPour',
                                            'textArchDownPour', 'textCirclePour', 'textButtonPour', 'textCurveUp',
                                            'textCurveDown', 'textCanUp', 'textCanDown', 'textWave1', 'textWave2',
                                            'textDoubleWave1', 'textWave4', 'textInflate', 'textDeflate',
                                            'textInflateBottom', 'textDeflateBottom', 'textInflateTop',
                                            'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate',
                                            'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown',
                                            'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown'])))
    avLst = Typed(expected_type=GeomGuideList, allow_none=True)

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class TextBodyProperties(Serialisable):

    rot = Typed(expected_type=Integer())
    spcFirstLastPara = Typed(expected_type=Bool, allow_none=True)
    vertOverflow = Typed(expected_type=Set(values=(['overflow', 'ellipsis', 'clip'])))
    horzOverflow = Typed(expected_type=Set(values=(['overflow', 'clip'])))
    vert = Typed(expected_type=Set(values=(['horz', 'vert', 'vert270',
                                            'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'])))
    wrap = Typed(expected_type=Set(values=(['none', 'square'])))
    lIns = Typed(expected_type=Coordinate, allow_none=True)
    tIns = Typed(expected_type=Coordinate, allow_none=True)
    rIns = Typed(expected_type=Coordinate, allow_none=True)
    bIns = Typed(expected_type=Coordinate, allow_none=True)
    numCol = Typed(expected_type=Integer())
    spcCol = Typed(expected_type=Coordinate)
    rtlCol = Typed(expected_type=Bool, allow_none=True)
    fromWordArt = Typed(expected_type=Bool, allow_none=True)
    anchor = Typed(expected_type=Set(values=(['t', 'ctr', 'b', 'just', 'dist'])))
    anchorCtr = Typed(expected_type=Bool, allow_none=True)
    forceAA = Typed(expected_type=Bool, allow_none=True)
    upright = Typed(expected_type=Bool, allow_none=True)
    compatLnSpc = Typed(expected_type=Bool, allow_none=True)
    prstTxWarp = Typed(expected_type=PresetTextShape, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 rot=None,
                 spcFirstLastPara=None,
                 vertOverflow=None,
                 horzOverflow=None,
                 vert=None,
                 wrap=None,
                 lIns=None,
                 tIns=None,
                 rIns=None,
                 bIns=None,
                 numCol=None,
                 spcCol=None,
                 rtlCol=None,
                 fromWordArt=None,
                 anchor=None,
                 anchorCtr=None,
                 forceAA=None,
                 upright=None,
                 compatLnSpc=None,
                 prstTxWarp=None,
                 scene3d=None,
                 extLst=None,
                ):
        self.rot = rot
        self.spcFirstLastPara = spcFirstLastPara
        self.vertOverflow = vertOverflow
        self.horzOverflow = horzOverflow
        self.vert = vert
        self.wrap = wrap
        self.lIns = lIns
        self.tIns = tIns
        self.rIns = rIns
        self.bIns = bIns
        self.numCol = numCol
        self.spcCol = spcCol
        self.rtlCol = rtlCol
        self.fromWordArt = fromWordArt
        self.anchor = anchor
        self.anchorCtr = anchorCtr
        self.forceAA = forceAA
        self.upright = upright
        self.compatLnSpc = compatLnSpc
        self.prstTxWarp = prstTxWarp
        self.scene3d = scene3d
        self.extLst = extLst


class TextBody(Serialisable):

    bodyPr = Typed(expected_type=TextBodyProperties, )
    lstStyle = Typed(expected_type=TextListStyle, allow_none=True)
    p = Typed(expected_type=TextParagraph, )

    def __init__(self,
                 bodyPr=None,
                 lstStyle=None,
                 p=None,
                ):
        self.bodyPr = bodyPr
        self.lstStyle = lstStyle
        self.p = p


class NumFmt(Serialisable):

    formatCode = String(allow_none=True)
    sourceLinked = Bool(allow_none=True)

    def __init__(self,
                 formatCode=None,
                 sourceLinked=None,
                ):
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked


class Tx(Serialisable):

    pass



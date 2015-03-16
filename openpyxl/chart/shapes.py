from openpyxl.descriptors.serialisable import Serialisable


from openpyxl.descriptors import (
    Set,
    )

class Shape(Serialisable):

    val = Set(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Point2D(Serialisable):

    x = Typed(expected_type=Coordinate, )
    y = Typed(expected_type=Coordinate, )

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class PositiveSize2D(Serialisable):

    cx = Typed(expected_type=Integer())
    cy = Typed(expected_type=Integer())

    def __init__(self,
                 cx=None,
                 cy=None,
                ):
        self.cx = cx
        self.cy = cy


class Transform2D(Serialisable):

    rot = Typed(expected_type=Integer())
    flipH = Typed(expected_type=Bool, allow_none=True)
    flipV = Typed(expected_type=Bool, allow_none=True)
    off = Typed(expected_type=Point2D, allow_none=True)
    ext = Typed(expected_type=PositiveSize2D, allow_none=True)

    def __init__(self,
                 rot=None,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext


class OfficeArtExtensionList(Serialisable):

    pass

class LineEndProperties(Serialisable):

    type = Typed(expected_type=Set(values=(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'])))
    w = Typed(expected_type=Set(values=(['sm', 'med', 'lg'])))
    len = Typed(expected_type=Set(values=(['sm', 'med', 'lg'])))

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class LineProperties(Serialisable):

    w = Typed(expected_type=Coordinate())
    cap = Typed(expected_type=Set(values=(['rnd', 'sq', 'flat'])))
    cmpd = Typed(expected_type=Set(values=(['sng', 'dbl', 'thickThin', 'thinThick', 'tri'])))
    algn = Typed(expected_type=Set(values=(['ctr', 'in'])))
    headEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    tailEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 headEnd=None,
                 tailEnd=None,
                 extLst=None,
                ):
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.algn = algn
        self.headEnd = headEnd
        self.tailEnd = tailEnd
        self.extLst = extLst


class SphereCoords(Serialisable):

    lat = Typed(expected_type=Integer)
    lon = Typed(expected_type=Integer)
    rev = Typed(expected_type=Integer)

    def __init__(self,
                 lat=None,
                 lon=None,
                 rev=None,
                ):
        self.lat = lat
        self.lon = lon
        self.rev = rev


class Camera(Serialisable):

    prst = Typed(expected_type=Set(values=(['legacyObliqueTopLeft', 'legacyObliqueTop', 'legacyObliqueTopRight', 'legacyObliqueLeft', 'legacyObliqueFront', 'legacyObliqueRight', 'legacyObliqueBottomLeft', 'legacyObliqueBottom', 'legacyObliqueBottomRight', 'legacyPerspectiveTopLeft', 'legacyPerspectiveTop', 'legacyPerspectiveTopRight', 'legacyPerspectiveLeft', 'legacyPerspectiveFront', 'legacyPerspectiveRight', 'legacyPerspectiveBottomLeft', 'legacyPerspectiveBottom', 'legacyPerspectiveBottomRight', 'orthographicFront', 'isometricTopUp', 'isometricTopDown', 'isometricBottomUp', 'isometricBottomDown', 'isometricLeftUp', 'isometricLeftDown', 'isometricRightUp', 'isometricRightDown', 'isometricOffAxis1Left', 'isometricOffAxis1Right', 'isometricOffAxis1Top', 'isometricOffAxis2Left', 'isometricOffAxis2Right', 'isometricOffAxis2Top', 'isometricOffAxis3Left', 'isometricOffAxis3Right', 'isometricOffAxis3Bottom', 'isometricOffAxis4Left', 'isometricOffAxis4Right', 'isometricOffAxis4Bottom', 'obliqueTopLeft', 'obliqueTop', 'obliqueTopRight', 'obliqueLeft', 'obliqueRight', 'obliqueBottomLeft', 'obliqueBottom', 'obliqueBottomRight', 'perspectiveFront', 'perspectiveLeft', 'perspectiveRight', 'perspectiveAbove', 'perspectiveBelow', 'perspectiveAboveLeftFacing', 'perspectiveAboveRightFacing', 'perspectiveContrastingLeftFacing', 'perspectiveContrastingRightFacing', 'perspectiveHeroicLeftFacing', 'perspectiveHeroicRightFacing', 'perspectiveHeroicExtremeLeftFacing', 'perspectiveHeroicExtremeRightFacing', 'perspectiveRelaxed', 'perspectiveRelaxedModerately'])))
    fov = Typed(expected_type=Integer)
    zoom = Typed(expected_type=Percentage, allow_none=True)
    rot = Typed(expected_type=SphereCoords, allow_none=True)

    def __init__(self,
                 prst=None,
                 fov=None,
                 zoom=None,
                 rot=None,
                ):
        self.prst = prst
        self.fov = fov
        self.zoom = zoom
        self.rot = rot


class LightRig(Serialisable):

    rig = Typed(expected_type=Set(values=(['legacyFlat1', 'legacyFlat2', 'legacyFlat3', 'legacyFlat4', 'legacyNormal1', 'legacyNormal2', 'legacyNormal3', 'legacyNormal4', 'legacyHarsh1', 'legacyHarsh2', 'legacyHarsh3', 'legacyHarsh4', 'threePt', 'balanced', 'soft', 'harsh', 'flood', 'contrasting', 'morning', 'sunrise', 'sunset', 'chilly', 'freezing', 'flat', 'twoPt', 'glow', 'brightRoom'])))
    dir = Typed(expected_type=Set(values=(['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br'])))
    rot = Typed(expected_type=SphereCoords, allow_none=True)

    def __init__(self,
                 rig=None,
                 dir=None,
                 rot=None,
                ):
        self.rig = rig
        self.dir = dir
        self.rot = rot


class Vector3D(Serialisable):

    dx = Typed(expected_type=Coordinate, )
    dy = Typed(expected_type=Coordinate, )
    dz = Typed(expected_type=Coordinate, )

    def __init__(self,
                 dx=None,
                 dy=None,
                 dz=None,
                ):
        self.dx = dx
        self.dy = dy
        self.dz = dz


class Point3D(Serialisable):

    x = Typed(expected_type=Coordinate, )
    y = Typed(expected_type=Coordinate, )
    z = Typed(expected_type=Coordinate, )

    def __init__(self,
                 x=None,
                 y=None,
                 z=None,
                ):
        self.x = x
        self.y = y
        self.z = z


class Backdrop(Serialisable):

    anchor = Typed(expected_type=Point3D, )
    norm = Typed(expected_type=Vector3D, )
    up = Typed(expected_type=Vector3D, )
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 anchor=None,
                 norm=None,
                 up=None,
                 extLst=None,
                ):
        self.anchor = anchor
        self.norm = norm
        self.up = up
        self.extLst = extLst


class Scene3D(Serialisable):

    camera = Typed(expected_type=Camera, )
    lightRig = Typed(expected_type=LightRig, )
    backdrop = Typed(expected_type=Backdrop, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 camera=None,
                 lightRig=None,
                 backdrop=None,
                 extLst=None,
                ):
        self.camera = camera
        self.lightRig = lightRig
        self.backdrop = backdrop
        self.extLst = extLst


class Bevel(Serialisable):

    w = Typed(expected_type=Integer())
    h = Typed(expected_type=Integer())
    prst = Typed(expected_type=Set(values=(['relaxedInset', 'circle', 'slope', 'cross', 'angle', 'softRound', 'convex', 'coolSlant', 'divot', 'riblet', 'hardEdge', 'artDeco'])))

    def __init__(self,
                 w=None,
                 h=None,
                 prst=None,
                ):
        self.w = w
        self.h = h
        self.prst = prst


class Shape3D(Serialisable):

    z = Typed(expected_type=Coordinate, allow_none=True)
    extrusionH = Typed(expected_type=Integer())
    contourW = Typed(expected_type=Integer())
    prstMaterial = Typed(expected_type=Set(values=(['legacyMatte', 'legacyPlastic', 'legacyMetal', 'legacyWireframe', 'matte', 'plastic', 'metal', 'warmMatte', 'translucentPowder', 'powder', 'dkEdge', 'softEdge', 'clear', 'flat', 'softmetal'])))
    bevelT = Typed(expected_type=Bevel, allow_none=True)
    bevelB = Typed(expected_type=Bevel, allow_none=True)
    extrusionClr = Typed(expected_type=Color, allow_none=True)
    contourClr = Typed(expected_type=Color, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 z=None,
                 extrusionH=None,
                 contourW=None,
                 prstMaterial=None,
                 bevelT=None,
                 bevelB=None,
                 extrusionClr=None,
                 contourClr=None,
                 extLst=None,
                ):
        self.z = z
        self.extrusionH = extrusionH
        self.contourW = contourW
        self.prstMaterial = prstMaterial
        self.bevelT = bevelT
        self.bevelB = bevelB
        self.extrusionClr = extrusionClr
        self.contourClr = contourClr
        self.extLst = extLst


class ShapeProperties(Serialisable):

    bwMode = Typed(expected_type=Set(values=(['clr', 'auto', 'gray', 'ltGray', 'invGray', 'grayWhite', 'blackGray', 'blackWhite', 'black', 'white', 'hidden'])))
    xfrm = Typed(expected_type=Transform2D, allow_none=True)
    ln = Typed(expected_type=LineProperties, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    sp3d = Typed(expected_type=Shape3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 ln=None,
                 scene3d=None,
                 sp3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.ln = ln
        self.scene3d = scene3d
        self.sp3d = sp3d
        self.extLst = extLst

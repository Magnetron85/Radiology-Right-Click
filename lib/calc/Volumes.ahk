; ============================================================
; lib/calc/Volumes.ahk -- Ellipsoid + Bullet volume
; ============================================================

#Requires AutoHotkey v2.0
#Include ..\Util.ahk

EllipsoidVolume_Entry(input) {
    return CalcEllipsoidVolume(input)
}

BulletVolume_Entry(input) {
    return CalcBulletVolume(input)
}

CalcEllipsoidVolume(input) {
    if !RegExMatch(input
        , "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
        , &m)
        return "Invalid input format for ellipsoid volume.`nExample: 3 x 2 x 1 cm"

    d := SortDimensions([m[1] + 0, m[2] + 0, m[3] + 0])
    isMm := (InStr(m[1], ".") = 0 && InStr(m[2], ".") = 0 && InStr(m[3], ".") = 0)
    if isMm
        d[1] /= 10, d[2] /= 10, d[3] /= 10

    volume  := (1/6) * 3.141592653589793 * d[1] * d[2] * d[3]
    rounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
    return input " (" rounded (Prefs.Get("display","units",true) ? " cc" : "") ")"
}

CalcBulletVolume(input) {
    if !RegExMatch(input
        , "\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*[x,]\s*(\d+(?:\.\d+)?)\s*"
        , &m)
        return "Invalid input format for bullet volume.`nExample: 3 x 2 x 1 cm"

    d := SortDimensions([m[1] + 0, m[2] + 0, m[3] + 0])
    isMm := (InStr(m[1], ".") = 0 && InStr(m[2], ".") = 0 && InStr(m[3], ".") = 0)
    if isMm
        d[1] /= 10, d[2] /= 10, d[3] /= 10

    volume  := d[1] * d[2] * d[3] * (5 * 3.141592653589793 / 24)
    rounded := (volume < 1) ? Round(volume, 3) : Round(volume, 1)
    return input " (" rounded (Prefs.Get("display","units",true) ? " cc" : "") ")"
}

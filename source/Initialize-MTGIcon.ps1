function Initialize-MTGIcon
{
    Add-Type -AssemblyName PresentationFramework

    $IconSource = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKUSURBVFhH7dZNiE5RHMfxx1tE8lKTl8JCFh
    OxoMaCbEQm2SiUhaQw    JYu7UVJWE7KQjVBsKLJQTFG6hYSSIpKF11CSMl7zPnx/zznnvhznMfc+9xkrpz517rn/e+9v7nPOuVOr3KJ4OGbZowFsUTwI07Ac23EC9/AdvbaqRS2Kx2ABunAAV/AWvxqoECCKZ2INutG
    Dp+hD6EGNVAoQumFZ/wOYAFE8ChMKaqtfQyd0w7JcgG3eeH/ONgrwCc9wGxdxGkewF6/h1zcboK9RgI76DUMtiu97teICTIaW7lc7Lo+wGB0ZO2DOJ5285gK4ZjarL3iHdjuatiheCnNt0skzAcyG1Ov5Ab/evYHNOI9l
    0A65CmNxAU/QY+sKB9DFofO+0Bw4Cl1/IzN2y9YNeIBXmIpr9ti5aetWJGNJJy+dA1E8BCMxDocQqvcDaFu/ip/Qq9eY+gsxCQ/sWMlJGMX7vDrHBZgHTTz3LdF8WQ+tDM0ffdTuIr02d5ByP4G+9VuxCetwDqF6F0Cvf
    T7e2HFRiA2Ygkt2LPXHgFFlDmjz0krQ+Clo49Lr3wj9nHuQfm2TTl7VSSgnsSRzrId22bqdyXjSyXMBRmA3tA2fgWa1/hq/3g+grXs0/E2r5DIMtSj+7NVKNoAePgz78QLH4OoKB9CE03IZXL8g2/4eoA1DsRJ65Qeh/y
    VnYw5m2Lp+Azj6qDzGZRzHLnyDX2cCqEXxXLzPnOu2Z9JmvhXmfNKpxr2BLQgF1Ep4mfER5lzSqcafhEU9bHWAdqwuSF/Kia0N0FQL37CsSgE6od9Os/wOsv9OFVUhgN/MJqJ1uxbaBfUBeo7Qg50WBmjUong8FkFL7TC
    u4wP+UYBQM7vcdHTakZKtVvsN9+yk2qyRCWkAAAAASUVORK5CYII="

    $Icon = New-Object System.Windows.Media.Imaging.BitmapImage
    $Icon.BeginInit()
    $Icon.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($IconSource)
    $Icon.EndInit()
    $Icon.Freeze()

    return $Icon
}

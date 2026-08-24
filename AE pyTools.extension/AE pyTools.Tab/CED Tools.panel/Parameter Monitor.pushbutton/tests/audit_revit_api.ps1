param(
    [Parameter(Mandatory = $true)]
    [string]$RevitApiPath
)

$ErrorActionPreference = "Stop"
$metadataContext = $null
try {
    $assembly = [System.Reflection.Assembly]::LoadFrom($RevitApiPath)
}
catch {
    $metadataAssemblyPath = Get-ChildItem (Split-Path -Parent $RevitApiPath) -Recurse -Filter "System.Reflection.MetadataLoadContext.dll" |
        Select-Object -First 1 -ExpandProperty FullName
    if (-not $metadataAssemblyPath) {
        throw
    }
    Add-Type -Path $metadataAssemblyPath
    $runtimePaths = [string][System.AppContext]::GetData("TRUSTED_PLATFORM_ASSEMBLIES") -split [System.IO.Path]::PathSeparator
    $assemblyPaths = [System.Collections.Generic.List[string]]::new()
    foreach ($runtimePath in $runtimePaths) {
        $assemblyPaths.Add([string]$runtimePath)
    }
    $assemblyPaths.Add([string]$RevitApiPath)
    $resolver = [System.Reflection.PathAssemblyResolver]::new($assemblyPaths)
    $metadataContext = [System.Reflection.MetadataLoadContext]::new($resolver)
    $assembly = $metadataContext.LoadFromAssemblyPath($RevitApiPath)
}

function Get-TypeContract {
    param([string]$FullName)
    $type = $assembly.GetType($FullName, $false)
    if ($null -eq $type) {
        return $null
    }
    return $type
}

function Get-MethodSignatures {
    param($Type, [string]$Name)
    if ($null -eq $Type) {
        return @()
    }
    return @(
        $Type.GetMethods() |
            Where-Object { $_.Name -eq $Name } |
            ForEach-Object {
                "{0}({1})" -f $_.Name, (($_.GetParameters() | ForEach-Object { $_.ParameterType.FullName }) -join ",")
            }
    )
}

$elementId = Get-TypeContract "Autodesk.Revit.DB.ElementId"
$element = Get-TypeContract "Autodesk.Revit.DB.Element"
$definition = Get-TypeContract "Autodesk.Revit.DB.Definition"
$labelUtils = Get-TypeContract "Autodesk.Revit.DB.LabelUtils"
$globalParameter = Get-TypeContract "Autodesk.Revit.DB.GlobalParameter"
$specString = Get-TypeContract "Autodesk.Revit.DB.SpecTypeId+String"
$dataStorage = Get-TypeContract "Autodesk.Revit.DB.ExtensibleStorage.DataStorage"
$linkInstance = Get-TypeContract "Autodesk.Revit.DB.RevitLinkInstance"

$contract = [ordered]@{
    assembly_version = $assembly.GetName().Version.ToString()
    element_id_value = $null -ne $elementId.GetProperty("Value")
    element_id_integer_value = $null -ne $elementId.GetProperty("IntegerValue")
    element_id_constructors = @(
        $elementId.GetConstructors() | ForEach-Object {
            (($_.GetParameters() | ForEach-Object { $_.ParameterType.FullName }) -join ",")
        }
    )
    element_get_parameters = Get-MethodSignatures $element "GetParameters"
    element_parameter_overloads = Get-MethodSignatures $element "get_Parameter"
    definition_get_data_type = (Get-MethodSignatures $definition "GetDataType")
    label_utils_get_label_for_spec = Get-MethodSignatures $labelUtils "GetLabelForSpec"
    global_parameter_create = Get-MethodSignatures $globalParameter "Create"
    multiline_text_spec = $null -ne $specString -and $null -ne $specString.GetProperty("MultilineText")
    extensible_storage_data_storage = $null -ne $dataStorage
    link_get_document = Get-MethodSignatures $linkInstance "GetLinkDocument"
    link_get_total_transform = Get-MethodSignatures $linkInstance "GetTotalTransform"
}

$contract | ConvertTo-Json -Depth 5 -Compress

if ($null -ne $metadataContext) {
    $metadataContext.Dispose()
}

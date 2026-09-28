Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Escape-Xml([string]$Value) {
  return [Security.SecurityElement]::Escape($Value)
}

function Para([string]$Text, [int]$Size = 21, [bool]$Bold = $false, [string]$Align = "left", [bool]$Bullet = $false, [int]$After = 120) {
  $boldTag = if ($Bold) { "<w:b/>" } else { "" }
  $numTag = if ($Bullet) { "<w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"1`"/></w:numPr>" } else { "" }
  $escaped = Escape-Xml $Text
  return "<w:p><w:pPr><w:jc w:val=`"$Align`"/>$numTag<w:spacing w:after=`"$After`"/></w:pPr><w:r><w:rPr>$boldTag<w:rFonts w:ascii=`"Microsoft YaHei`" w:eastAsia=`"Microsoft YaHei`"/><w:sz w:val=`"$Size`"/></w:rPr><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>"
}

$itemsOne = @(
  "使用 text、password、email、tel、date、checkbox 等 input 类型，分别接收用户名、密码、邮箱、手机号、出生日期和协议确认信息。",
  "使用 required 属性实现必填校验，未填写完整时浏览器会自动提示。",
  "使用 pattern 正则表达式校验用户名、邮箱和中国大陆手机号码格式。",
  "使用 minlength、maxlength 限制密码长度为 6 至 20 位。",
  "使用 placeholder 为输入框提供填写提示。",
  "使用 autocomplete 支持浏览器自动填充账号、密码、邮箱、手机号和出生日期。",
  "登录页面使用 autofocus，打开页面时自动聚焦用户名输入框。",
  "使用 label 的 for 属性与输入框 id 绑定，提升表单可用性。",
  "使用 fieldset 和 legend 对登录、注册表单进行分组，使结构更清晰。"
)
$itemsTwo = @(
  "使用 header 放置平台图标、页面标题和副标题。",
  "使用 section 放置登录或注册表单卡片。",
  "使用 footer 放置前往注册和去登录的页面跳转链接。",
  "使用 meta viewport 实现移动端自适应。",
  "使用 CSS3 的渐变背景、圆角、阴影、弹性布局和媒体查询，实现统一、简约的页面视觉效果。",
  "登录页与注册页通过超链接相互跳转，形成基本的用户入口流程。"
)

$body = ""
$body += Para "AI助教答疑平台作业说明" 40 $true "center" $false 280
$body += Para "项目内容：使用原生 HTML5 和 CSS 完成 AI 助教答疑平台的登录、注册页面。页面采用淡蓝渐变背景、居中白色圆角卡片和蓝色按钮，并适配移动端显示。" 21 $false "left" $false 200
$body += Para "一、用到的 HTML5 表单增强特性" 28 $true "left" $false 140
foreach ($item in $itemsOne) { $body += Para $item 21 $false "left" $true 70 }
$body += Para "二、用到的其他 HTML5 相关能力" 28 $true "left" $false 140
foreach ($item in $itemsTwo) { $body += Para $item 21 $false "left" $true 70 }

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>$body<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1300" w:right="1440" w:bottom="1300" w:left="1440"/></w:sectPr></w:body></w:document>
"@
$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>
"@
$rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
"@
$docRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>
"@
$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:lang w:val="zh-CN" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>
"@
$numbering = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>
"@

$target = Join-Path (Get-Location) "AI助教答疑平台作业说明.docx"
$temp = Join-Path $env:TEMP ("ai-tutor-docx-" + [guid]::NewGuid())
New-Item -ItemType Directory -Path "$temp\_rels", "$temp\word\_rels" -Force | Out-Null
[IO.File]::WriteAllText("$temp\[Content_Types].xml", $contentTypes, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText("$temp\_rels\.rels", $rels, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText("$temp\word\document.xml", $documentXml, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText("$temp\word\styles.xml", $styles, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText("$temp\word\numbering.xml", $numbering, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText("$temp\word\_rels\document.xml.rels", $docRels, [Text.UTF8Encoding]::new($false))
if (Test-Path $target) { Remove-Item -LiteralPath $target -Force }
$archive = [IO.Compression.ZipFile]::Open($target, [IO.Compression.ZipArchiveMode]::Create)
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\[Content_Types].xml", "[Content_Types].xml") | Out-Null
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\_rels\.rels", "_rels/.rels") | Out-Null
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\word\document.xml", "word/document.xml") | Out-Null
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\word\styles.xml", "word/styles.xml") | Out-Null
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\word\numbering.xml", "word/numbering.xml") | Out-Null
[IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, "$temp\word\_rels\document.xml.rels", "word/_rels/document.xml.rels") | Out-Null
$archive.Dispose()
Remove-Item -LiteralPath $temp -Recurse -Force
Write-Output $target

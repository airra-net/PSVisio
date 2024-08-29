<#
.SYNOPSIS
    Microsoft Powershell functions for operate Visio Drawing

.DESCRIPTION
    Microsoft Powershell functions for drawing Visio objects: Document, Page, Stensil, etc.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development
                    Create function New-VisioApplication.

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Begin Reorganize script.

    Version:        3.3
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  28.06.2023
    Purpose/Change: Reorganize Draw-VisioItem Function. Add parameter LineWeight.

    Version:        3.4
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  22.08.2024
    Purpose/Change: Reorganize Draw-VisioItem Function. Add parameter LineColor.

    Version:        3.5
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  29.08.2024
    Purpose/Change: Added the previously lost function Resize-VisioPageToFitContents.

    Version:        3.6
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  29.08.2024
    Purpose/Change: Added the previously lost function Save-VisioDocument.

    Version:        3.7
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  29.08.2024
    Purpose/Change: Added the previously lost function Close-VisioApplication.
   ...
   
.EXAMPLE

    Load Microsoft Powershell functions for operate Visio Drawing: 

    . .\PSVisio.ps1 
#>

# Set Variables
$Shape = 0
$Line = 0
$Icon = 0

Function New-VisioApplication {

<#
.SYNOPSIS
    Microsoft Powershell function for create Visio Application

.DESCRIPTION
    Microsoft Powershell function for create Visio Application.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    New-VisioApplication 
#>

# Create Visio Object
$Script:Application = New-Object -ComObject Visio.Application
$Script:Application.Visible = $True

}

Function New-VisioDocument {

<#
.SYNOPSIS
    Microsoft Powershell function for create Visio Document

.DESCRIPTION
    Microsoft Powershell function for create Visio Document.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    New-VisioDocument 
#>

# Create Document from Blank Template
$Script:Documents = $Script:Application.Documents
$Script:Document = $Script:Application.Documents.Add('')

}

Function Set-VisioPage {

<#
.SYNOPSIS
    Microsoft Powershell function for create Visio Document Page

.DESCRIPTION
    Microsoft Powershell function for create Visio Document Page.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Set-VisioPage 
#>

# Set Visio Active Page
$Script:Page = $Script:Application.ActivePage
$Script:Application.ActivePage.PageSheet

}

Function Add-VisioStensil  {

<#
.SYNOPSIS
    Microsoft Powershell function for Add Visio Stensil

.DESCRIPTION
    Microsoft Powershell function for Add Visio Stensil.

.PARAMETER Name
    Name Identifier of Visio Stensils.

.PARAMETER File
    Name of Visio Stensils file.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Add-VisioStensil -Name "Basic" -File "BASIC_M.vss" 
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$Name,
        
    [Parameter(Mandatory)]
    [String]$File         
)

# Set Expression and Add Visio Stensil
$Expression = '$Script:' + $Name + ' = $Script:Application.Documents.Add("' + $File +'")'
Invoke-Expression $Expression

}

Function Set-VisioStensilMasterItem {

<#
.SYNOPSIS
    Microsoft Powershell function for Set Visio Stensil Master Item

.DESCRIPTION
    Microsoft Powershell function for Set Visio Stensil Master Item

.PARAMETER Stensil
    Name Identifier of pre-added Visio Stensils.

.PARAMETER Item
    Reference Name Identifier of Visio Stensils Item .
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Set-VisioStensilMasterItem -Stensil "Basic" -Item "Rectangle" 
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$Stensil,
        
    [Parameter(Mandatory)]
    [String]$Item     
)

# Set Expression And Set Masters Item Rectangle
$ItemWithoutSpace = $Item -replace " ",""
$Expression = '$Script:' + $ItemWithoutSpace + ' = $Script:' + $Stensil + '.Masters.Item("' + $Item + '")'
Invoke-Expression $Expression

}

Function Draw-VisioItem {

<#
.SYNOPSIS
    Microsoft Powershell function for Draw Visio Item

.DESCRIPTION
    Microsoft Powershell function for Draw Visio Item.

.PARAMETER Master
    Name Identifier of Master Item Visio Stensils.

.PARAMETER X
    X coordinate of Visio Stensils Item.

.PARAMETER Y
    Y coordinate of Visio Stensils Item.

.PARAMETER Width
    Width size of Visio Stensils Item.

.PARAMETER Height
    Height size of Visio Stensils Item.

.PARAMETER FillForegnd
    Foreground color of Visio Stensils Item.

.PARAMETER Fill
    Background color of Visio Stensils Item.

.PARAMETER LinePattern
    Contour line style of Visio Stensils Item.

.PARAMETER LineWeight
    Contour Line thickness size.

.PARAMETER LineColor
    Contour Line Color.

.PARAMETER Text
    Text Visio Stensils Item.

.PARAMETER VerticalAlign
    Vertical Align Visio Stensils Item.

.PARAMETER ParaHorzAlign
    Horizontal Align Visio Stensils Item.

.PARAMETER CharSize
    Text Character Size of Visio Stensils Item.

.PARAMETER CharColor
    Text Character Color of Visio Stensils Item.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  14.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function

    Version:        3.3
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  28.06.2023
    Purpose/Change: Reorganize function. Add parameter LineWeight.

    Version:        3.4
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  22.08.2024
    Purpose/Change: Reorganize function. Add parameter LineColor.    
   
.EXAMPLE

    Run:

    Draw-VisioItem -Master "Rectangle" -X 6.375 -Y 7.125 -Width 12.2501 -Height 7.25 -FillForegnd "RGB(0,153,204)"`
    -LinePattern 0 -LineWeight "1 pt" -Text "Microsoft Virtual Machine Manager Architecture" -VerticalAlign 0 -ParaHorzAlign 0`
    -CharSize "20 pt" -CharColor "RGB(255,255,255)" -Fill "RGB(255,255,255)" -LineColor "RGB(255,255,255)"
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$Master,
        
    [Parameter(Mandatory)]
    [String]$X,

    [Parameter(Mandatory)]
    [String]$Y,

    [Parameter()]
    [String]$Width,

    [Parameter()]
    [String]$Height,

    [Parameter()]
    [String]$FillForegnd,

    [Parameter()]
    [String]$Fill,

    [Parameter()]
    [String]$LinePattern,

    [Parameter()]
    [String]$LineWeight,

    [Parameter()]
    [String]$LineColor,

    [Parameter()]
    [String]$Text,

    [Parameter()]
    [String]$VerticalAlign,

    [Parameter()]
    [String]$ParaHorzAlign,

    [Parameter()]
    [String]$CharSize,
    
    [Parameter()]
    [String]$CharColor        
)

# Set Variables
$Script:Shape++
$Master = $Master -replace " ",""

# Set Expression And Draw Item
$Expression = '$Script:Shape' + $Script:Shape + ' = $Script:Page.Drop(' + '$' + $Master + ',' + $X + ',' + $Y + ')'
Invoke-Expression $Expression

# Set Item Width Properties
If ($Width)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Width").Formula = ' + $Width
		Invoke-Expression $Expression
	}

# Set Item Height Properties
If ($Height)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Height").Formula = ' + $Height
		Invoke-Expression $Expression
	}

# Set Item FillForegnd Properties
If ($FillForegnd)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("FillForegnd").FormulaU = "=' +  $FillForegnd + '"'
		Invoke-Expression $Expression
	}

# Set Item Fill Properties
If ($Fill)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.CellsU("FillForegnd").FormulaForceU = "' +  $Fill + '"'
		Invoke-Expression $Expression
	}

# Set Item LinePattern Properties
If ($LinePattern)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("LinePattern").Formula = ' + $LinePattern
		Invoke-Expression $Expression
	}

# Set Item LineWeight Properties
If ($LineWeight)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("LineWeight").Formula = "' + $LineWeight + '"'
		Invoke-Expression $Expression
	}

# Set Item Line Color Properties
If ($LineColor)
	{
        $Expression = '$Script:Shape' + $Script:Shape + '.Cells("LineColor").FormulaU = "=' +  $LineColor + '"'
		Invoke-Expression $Expression
	}

# Set Item Text
If ($Text)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Item VerticalAlign Properties
If ($VerticalAlign)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("VerticalAlign").Formula = ' + $VerticalAlign
		Invoke-Expression $Expression
	}

# Set Item HorzAlign Properties
If ($ParaHorzAlign)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Para.HorzAlign").Formula = ' + $ParaHorzAlign
		Invoke-Expression $Expression
	}

# Set Item Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}

# Set Item Char.Color Properties
If ($CharColor)
	{
		$Expression = '$Script:Shape' + $Script:Shape + '.Cells("Char.Color").FormulaU = "=' +  $CharColor + '"'
		Invoke-Expression $Expression
	}
}

Function Draw-VisioLine {

<#
.SYNOPSIS
    Microsoft Powershell function for Draw Visio Line

.DESCRIPTION
    Microsoft Powershell function for Draw Visio Line.

.PARAMETER BeginX
    Begin X coordinate of Visio Line.

.PARAMETER BeginY
    Begin Y coordinate of Visio Line.

.PARAMETER EndX
    End X coordinate of Visio Line.

.PARAMETER EndY
    End X coordinate of Visio Line.

.PARAMETER LineWeight
    Visio Line thickness size.

.PARAMETER LineColor
    Visio Line Color.
    
.PARAMETER BeginArrow
    Begin Arrow Visio Line Style.

.PARAMETER EndArrow
    End Arrow Visio Line Style.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  15.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Draw-VisioLine -BeginX 0.3125 -BeginY 10.3438 -EndX 12.4948 -EndY 10.3438 -LineWeight "1 pt"`
    -LineColor "RGB(255,255,255)" -BeginArrow 4 -EndArrow 4 
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$BeginX,
        
    [Parameter(Mandatory)]
    [String]$BeginY,

    [Parameter(Mandatory)]
    [String]$EndX,

    [Parameter(Mandatory)]
    [String]$EndY,

    [Parameter()]
    [String]$LineWeight,

    [Parameter()]
    [String]$LineColor,

    [Parameter()]
    [String]$BeginArrow,

    [Parameter()]
    [String]$EndArrow      
)

# Set variable
$Script:Line++

# Set Expression And Draw Line
$Expression = '$Script:Line' + $Script:Line + ' = $Script:Page.DrawLine(' + $BeginX + ',' + $BeginY + ',' + $EndX + ',' + $EndY + ')'
Invoke-Expression $Expression

# Set Line Width Properties
If ($LineWeight)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("LineWeight").Formula = "' + $LineWeight + '"'
		Invoke-Expression $Expression
	}

# Set Line Color Properties
$Expression = '$Script:Line' + $Script:Line + '.Cells("LineColor").FormulaU = "=' +  $LineColor + '"'
Invoke-Expression $Expression

# Set Line Begin Arrow Properties
If ($BeginArrow)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("BeginArrow").Formula = ' + $BeginArrow
		Invoke-Expression $Expression
	}

# Set Line End Arrow Properties
If ($EndArrow)
	{
		$Expression = '$Script:Line' + $Script:Line + '.Cells("EndArrow").Formula = ' + $EndArrow
		Invoke-Expression $Expression
	}
}

Function Draw-VisioIcon {

<#
.SYNOPSIS
    Microsoft Powershell function for Draw Visio Icon

.DESCRIPTION
    Microsoft Powershell function for Draw Visio Icon.

.PARAMETER IconPath
    Path to load icon.

.PARAMETER Width
    Width size of Visio Icon.

.PARAMETER Height
    Height size of Visio Icon.

.PARAMETER PinX
    X coordinates of Visio Icon.

.PARAMETER PinY
    Y Coordinates of Visio Icon.

.PARAMETER Text
    Text Visio Icon.

.PARAMETER CharSize
    Text Character Size of Visio Icon.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  16.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  02.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Draw-VisioIcon -IconPath "c:\!\powershell.png" -Width 0.9843 -Height 0.9843 -PinX 2.5547 -PinY 9.2682`
    -Text "Windows Powershell" -CharSize "10 pt"
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$IconPath,

    [Parameter(Mandatory)]
    [String]$Width,

    [Parameter(Mandatory)]
    [String]$Height,

    [Parameter()]
    [String]$PinX,

    [Parameter()]
    [String]$PinY,

    [Parameter()]
    [String]$Text,

    [Parameter()]
    [String]$CharSize
)

# Set Variables
$Script:Icon++

# Import Icon Item
$Expression = '$Script:Icon' + $Script:Icon + ' = $Script:Page.Import("' + $IconPath + '")'
Invoke-Expression $Expression

# Set Icon Width Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Width").Formula = ' + $Width
Invoke-Expression $Expression

# Set Icon Height Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Height").Formula = ' + $Height
Invoke-Expression $Expression

# Set Icon PinX Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("PinX").Formula = ' + $PinX
Invoke-Expression $Expression

# Set Icon PinY Properties
$Expression = '$Script:Icon' + $Script:Icon + '.Cells("PinY").Formula = ' + $PinY
Invoke-Expression $Expression

# Set Icon Text
If ($Text)
	{
		$Expression = '$Script:Icon' + $Script:Icon + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Icon Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Icon' + $Script:Icon + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}
}

Function Draw-VisioText {

<#
.SYNOPSIS
    Microsoft Powershell function for Draw Visio Text

.DESCRIPTION
    Microsoft Powershell function for Draw Visio Text.

.PARAMETER BeginX
    Begin X coordinate of Visio Text.

.PARAMETER BeginY
    Begin Y coordinate of Visio Text.

.PARAMETER Width
    Width size of Visio Text.

.PARAMETER Height
    Height size of Visio Text.

.PARAMETER FillForegnd
    Background color of Visio Text.

.PARAMETER LinePattern
    Contour line style of Visio Text.

.PARAMETER Text
    Text Visio Item.

.PARAMETER VerticalAlign
    Vertical Align Visio Text.

.PARAMETER ParaHorzAlign
    Horizontal Align Visio Text.

.PARAMETER CharSize
    Text Character Size of Visio Text.

.PARAMETER CharColor
    Text Character Color of Visio Text.

.PARAMETER CharStyle
    Text Character Style of Visio Text.

.PARAMETER FillForegndTrans
    Transparent Fill Value of Visio Text.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  15.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  02.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Draw-VisioText -X 4.25 -Y 8.875 -Width 1.3751 -Height 0.375 -Text "Deploy Admin / Dev" -CharSize "10 pt" -CharStyle 17 -LinePattern "0" -FillForegndTrans "100%"
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$X,
        
    [Parameter(Mandatory)]
    [String]$Y,

    [Parameter(Mandatory)]
    [String]$Width,

    [Parameter(Mandatory)]
    [String]$Height,

    [Parameter()]
    [String]$FillForegnd,

    [Parameter()]
    [String]$LinePattern,

    [Parameter()]
    [String]$Text,

    [Parameter()]
    [String]$VerticalAlign,

    [Parameter()]
    [String]$ParaHorzAlign,

    [Parameter()]
    [String]$CharSize,

    [Parameter()]
    [String]$CharColor,

    [Parameter()]
    [String]$CharStyle,

    [Parameter()]
    [String]$FillForegndTrans
)

# Set Variables
$Script:Text++
$Master = "Rectangle"

# Set Expression And Draw Text
$Expression = '$Script:Text' + $Script:Text + ' = $Script:Page.Drop(' + '$' + $Master + ',' + $X + ',' + $Y + ')'
Invoke-Expression $Expression

# Set Item Width Properties
If ($Width)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Width").Formula = ' + $Width
		Invoke-Expression $Expression
	}

# Set Item Height Properties
If ($Height)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Height").Formula = ' + $Height
		Invoke-Expression $Expression
	}

# Set Item FillForegnd Properties
If ($FillForegnd)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("FillForegnd").Formula = "=' +  $FillForegnd + '"'
		Invoke-Expression $Expression
	}

# Set Item LinePattern Properties
If ($LinePattern)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("LinePattern").Formula = ' + $LinePattern
		Invoke-Expression $Expression
	}

# Set Item Text
If ($Text)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Text = "' + $Text + '"'
		Invoke-Expression $Expression
	}

# Set Item VerticalAlign Properties
If ($VerticalAlign)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("VerticalAlign").Formula = ' + $VerticalAlign
		Invoke-Expression $Expression
	}

# Set Item HorzAlign Properties
If ($ParaHorzAlign)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Para.HorzAlign").Formula = ' + $ParaHorzAlign
		Invoke-Expression $Expression
	}

# Set Item Char.Size Properties
If ($CharSize)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Size").Formula = "' + $CharSize + '"'
		Invoke-Expression $Expression
	}

# Set Item Char.Color Properties
If ($CharColor)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Color").FormulaU = "=' +  $CharColor + '"'
		Invoke-Expression $Expression
	}

# Set Item Char.Style Properties
If ($CharStyle)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("Char.Style").Formula = "' + $CharStyle + '"'
		Invoke-Expression $Expression
	}
	
# Set Item FillForegndTrans Properties
If ($FillForegndTrans)
	{
		$Expression = '$Script:Text' + $Script:Text + '.Cells("FillForegndTrans").Formula = "' + $FillForegndTrans + '"'
		Invoke-Expression $Expression
	}		
}

Function Draw-VisioPolyLine {

<#
.SYNOPSIS
    Microsoft Powershell function for Draw Visio PolyLine

.DESCRIPTION
    Microsoft Powershell function for Draw Visio PolyLine.

.PARAMETER Polyline
   Polyline coordinates of Visio PolyLine.

.PARAMETER LineWeight
    Visio Line thickness size.

.PARAMETER LineColor
    Visio Line Color.
    
.PARAMETER BeginArrow
    Begin Arrow Visio Line Style.

.PARAMETER EndArrow
    End Arrow Visio Line Style.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  18.09.2007
    Purpose/Change: Initial script development

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function
   
.EXAMPLE

    Run:

    Draw-VisioPolyLine -Polyline 1.0938,9.0625,1.4063,9.0625,1.4063,8.6563,1.0938,8.6563 -LineWeight "0.5 pt"`
    -LineColor "RGB(255,255,255)" -BeginArrow 1 -EndArrow 1
#>

Param ( 
    [Parameter(Mandatory)]
    [array]$Polyline,

    [Parameter()]
    [String]$LineWeight,

    [Parameter()]
    [String]$LineColor,

    [Parameter()]
    [String]$BeginArrow,

    [Parameter()]
    [String]$EndArrow      
)

# Set  Variable
$Script:PolyLine++
[double[]]$PolyLineCoordinates=@()
$PolyLineCoordinates += $Polyline

# Set Expression And Draw PolyLine
$Expression = '$Script:PolyLine' + $Script:PolyLine + ' = $Script:Page.DrawPolyLine([ref]($PolyLineCoordinates),0)'
Invoke-Expression $Expression

# Set Line Width Properties
If ($LineWeight)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("LineWeight").Formula = "' + $LineWeight + '"'
		Invoke-Expression $Expression
	}

# Set Line Color Properties
$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("LineColor").FormulaU = "=' +  $LineColor + '"'
Invoke-Expression $Expression

# Set Line Begin Arrow Properties
If ($BeginArrow)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("BeginArrow").Formula = ' + $BeginArrow
		Invoke-Expression $Expression
	}

# Set Line End Arrow Properties
If ($EndArrow)
	{
		$Expression = '$Script:PolyLine' + $Script:PolyLine + '.Cells("EndArrow").Formula = ' + $EndArrow
		Invoke-Expression $Expression
	}
}

Function Resize-VisioPageToFitContents {
<#
.SYNOPSIS
    Microsoft Powershell function for Resize Active Visio Document Page to Fit Contents.

.DESCRIPTION
    Microsoft Powershell function for Resize Active Visio Document Page to Fit Contents.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  17.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function.
   
.EXAMPLE

    Run:

    Resize-VisioPageToFitContents 
#>

# Resize Page to Fit Contents
$Script:Page.ResizeToFitContents()

}

Function Save-VisioDocument {
<#
.SYNOPSIS
    Microsoft Powershell function for save Visio Document.

.DESCRIPTION
    Microsoft Powershell function for save Visio Document.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  17.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function.
   
.EXAMPLE

    Run:

    Save-VisioDocument -File 'C:\!\Diagram.vsd' 
#>

Param ( 
    [Parameter(Mandatory)]
    [String]$File
)

# Save Document
$Expression = '$Script:Document.SaveAs("' + $File + '")'
Invoke-Expression $Expression

}

Function Close-VisioApplication {
<#
.SYNOPSIS
    Microsoft Powershell function for Close Visio Application.

.DESCRIPTION
    Microsoft Powershell function for Close Visio Application.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  17.09.2007
    Purpose/Change: Initial script development

...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  01.09.2022
    Purpose/Change: Reorganize function.
   
.EXAMPLE

    Run:

    Close-VisioApplication 
#>

# Close Visio Application
$Script:Application.Quit()

}

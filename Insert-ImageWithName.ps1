<#
.SYNOPSIS
    Insert image printouts with filenames in OneNote.

.DESCRIPTION
    This is a OneMore plugin that automates insertion of image printouts along with their respective filenames.
    Its purpose is to increase data traceability and to avoid unnecessary bloating of notebook file sizes.
    For consistency's sake, these printouts have a fixed width, which can be altered by changing the $NewWidth variable.

.PARAMETER Path
    The path of a OneNote page XML file.
    The plugin must update this file in order for changes to be applied by OneMore.
    If no changes are detected then the current page is not updated.

.NOTES
    Make sure to change the plugin timeout to 0 in the OneMore Settings to give you enough time to select the files.
    This version does not include any file type limitations or error handling.

    Adapted from 'Get-MeetingDate.ps1' by stevencohn
    and 'PowerShell: Resize-Image' by someshinyobject (Christopher Walker)
    Written by:
    Joris Van Meenen
#>

[CmdletBinding(SupportsShouldProcess = $true)]

param (
	[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
	[string] $Path
)

Begin
{
	function UpdatePageXml ($filePath)
	{
		#Load dependencies
        $null = [Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq")
        Add-Type -AssemblyName System.Windows.Forms   
        Add-Type -AssemblyName System.Drawing    

        #Load XML
		$xml = [Xml.Linq.XElement]::Load($filePath)

		Write-Host "Loaded $filepath"
        
        #Get NameSpace
        $ns = $xml.GetNamespaceOfPrefix('one')

        #Find selection
        $OENode = $xml.Descendants($ns + "Outline").Descendants($ns + "OEChildren").Descendants($ns + "OE")
        $SNode = $OENode | Where-Object {$_.Descendants().Attribute("selected").Value -eq 'all'}

        #Open File Selection Dialog
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
        $FileBrowser.MultiSelect = $true
        $null = $FileBrowser.ShowDialog()
        $Paths = $FileBrowser.FileNames
        Write-Host $Paths
        [array]::Reverse($Paths)
        Write-Host $Paths
        $FileNames = $FileBrowser.SafeFileNames
        Write-Host $FileNames
        [array]::Reverse($FileNames)
        Write-Host $FileNames

        ForEach ($Path in $Paths) {
            
            $Index = $Paths.IndexOf($Path)
            Write-Host $Path
            Write-Host $Index

            #Load and transform image
            $OldImage = new-object System.Drawing.Bitmap $Path
            $OldHeight = $OldImage.height
            $OldWidth = $OldImage.width

            $NewWidth = 360
            $NewHeight = $OldHeight / $OldWidth * $NewWidth
            $Bitmap = New-Object System.Drawing.Bitmap $NewWidth, $NewHeight
            $NewImage = [System.Drawing.Graphics]::FromImage($Bitmap)
            $NewImage.DrawImage($OldImage, $(New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $NewWidth, $NewHeight))

            #Convert new image to uri
            $memory = New-Object System.IO.MemoryStream
            $null = $Bitmap.Save($memory, "PNG")
            [byte[]]$bytes = $memory.ToArray()
            $uri = [System.Convert]::ToBase64String($bytes)

            #Clone selection node to get all appropriate references
            $OEClone = New-Object Xml.Linq.XElement $SNode
            $OEClone2 = New-Object Xml.Linq.XElement $SNode

            #Delete children
            $OEClone.RemoveNodes()
            $OEClone2.RemoveNodes()

            #Construct Image node
            $ImageNode = New-Object System.Xml.Linq.XElement($ns + "Image")
            $ImageNode.SetAttributeValue('format','png')

            #Construct Size node
            $SizeNode = New-Object Xml.Linq.XElement($ns + "Size")
            $SizeNode.SetAttributeValue('width',$NewWidth)
            $SizeNode.SetAttributeValue('height',$NewHeight)
            $SizeNode.SetAttributeValue('isSetByUser','true')

            #Construct Data node
            $DataNode = New-Object Xml.Linq.XElement($ns + "Data")
            $DataNode.Value = $uri

            #Construct Text node
            $TextNode = New-Object System.Xml.Linq.XElement($ns + "T")

            #Construct CData node
            $CDataNode = New-Object System.Xml.Linq.XCData($FileNames[$Index])

            #Add new nodes together
            $ImageNode.Add($SizeNode)
            $ImageNode.Add($DataNode)
            $OEClone.Add($ImageNode)

            $TextNode.Add($CDataNode)
            $OEClone2.Add($TextNode)
        
            #Insert new nodes below selection
            $SNode.AddAfterSelf($OEClone)
            $SNode.AddAfterSelf($OEClone2)

        }

        #Save modified XML
		$xml.Save($filePath, [Xml.Linq.SaveOptions]::None)
		Write-Host "Saved $filepath"
	}
}
Process
{
	$filepath = Resolve-Path $Path -ErrorAction SilentlyContinue
	if (!$filepath)
	{
		Write-Host "Could not find file $Path" -ForegroundColor Yellow
		return
	}

    UpdatePageXml $filepath

    Write-Host 'Done'
}
# Switch Statment to output the groups to the console
function Show-Groups ($sg_counter) {

    $msg = -join ((' ' * (3 *($sg_counter - 1))), $groupName)

    switch ($sg_counter) {
        3 {Write-Host -ForegroundColor Green -Object $msg ;break}
        4 {Write-Host -ForegroundColor Cyan -Object $msg ;break}
        5 {Write-Host -ForegroundColor Magenta -Object $msg ;break}
        6 {Write-Host -ForegroundColor Gray -Object $msg ;break}
        7 {Write-Host -ForegroundColor red -Object $msg ;break}
        8 {Write-Host -ForegroundColor darkyellow -Object $msg ;break}
        9 {Write-Host -ForegroundColor DarkGreen -Object $msg ;break}
        10 {Write-Host -ForegroundColor Darkcyan -Object $msg ;break}
        11 {Write-Host -ForegroundColor DarkMagenta -Object $msg ;break}
        12 {Write-Host -ForegroundColor DarkGray -Object $msg ;break}
        13 {Write-Host -ForegroundColor DarkRed -Object $msg ;break}
        14 {Write-Host -ForegroundColor White -Object $msg ;break}
        15 {Write-Host -ForegroundColor black -Object $msg ;break }
        16 {Write-Host -ForegroundColor darkyellow -Object $msg ;break}
        17..30 {Write-Host "need more swiches"}
        default {Write-Host -foregroundcolor Yellow -Object $msg ;break}
    }
}


# Function to increment the script level variable rowcounter
function Increment-Rowcounter {

    $Script:rowcounter++

}


# Switch statement to put items into the Excel cells, color the cells, and output to the console
function Export-ToExcel ($ETE_counter, $ETE_rowcounter) {
    switch ($ETE_counter) {

        3 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 4
            Increment-Rowcounter
                
        }

        4 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 8
            Increment-Rowcounter
                
        }

        5 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 18
            Increment-Rowcounter

        }

        6 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 16
            Increment-Rowcounter
        }

        7 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 3
            Increment-Rowcounter
        }

        8 {

            Show-Groups $ETE_counter   
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 45
            Increment-Rowcounter

        }

        9 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 10
            Increment-Rowcounter

        }

        10 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 33
            Increment-Rowcounter

        }

        11 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 29
            Increment-Rowcounter

        }

        12 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 56
            Increment-Rowcounter

        }

        13 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 30
            Increment-Rowcounter

        }

        14 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 2
            Increment-Rowcounter

        }

        15 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 51
            Increment-Rowcounter

        }

        16 {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 45
            Increment-Rowcounter

        }

        17..30 {
                
            Write-Host "need more swiches"
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            Increment-Rowcounter
        
        }

        default {

            Show-Groups $ETE_counter
            $ws.cells.item($ETE_rowcounter, $ETE_counter) = $groupName
            $ws.cells.item($ETE_rowcounter, $ETE_counter).interior.colorIndex = 6
            Increment-Rowcounter
            
        }
    }
}


# Function that runs against each group in the pipeline that gathers its Nested Groups and then calls itself against the new Nested Groups in the pipeline until there are no nested groups
function Find-NestedGroups ($FNG_counter, $FNG_rowcounter){
    
    $groupName = $_.name
    $objectGUID = $_.objectGUID
    $group = get-adgroupmember $_.objectGUID | where {$_.objectClass -like "Group"} 
            
    try {
        $childGroup = get-adgroupmember $group | where {$_.objectClass -like "Group"}
    }
    catch {}
    
    if ($WithExcel) {
    
        Export-ToExcel $FNG_counter $SCript:rowcounter
        
    }
    else {
    
        Show-Groups $FNG_counter
           
    }
        
    # Prevents an infinite loop of nested groups by checking to see if the 'child' group is also the 'parent' of the current group
    if ($objectGUID -ccontains $childGroup.objectGUID ) {
    
        $FNG_counter++
            
        $groupName = (Get-ADGroupMember $objectGUID | where {$_.objectClass -like "Group"}).name
    
        if ($WithExcel) {
            Export-ToExcel $FNG_counter $Script:rowcounter
        }
        else {
            Show-Groups $FNG_counter
        }
    }
    else {
        
        $FNG_counter++
        $group | ForEach-Object {Find-NestedGroups $FNG_counter $Script:rowcounter}
        
            
    }
}
    

function Get-NestedGroups {

    [CmdletBinding()]
    Param(
        
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true)]
        [String]$TopLevelGroup,

        [Parameter(ParameterSetName = 'Excel', Mandatory = $false)]
        [Switch]$WithExcel,

        [parameter(ParameterSetName = 'Excel', Mandatory = $true)]
        [String]$ExcelFilePath

    ) 

    # If outputting to Excel, starts excel and opens or creates the file to be used.
    if ($WithExcel) {

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true

        if (Test-Path $ExcelFilePath) {
            $xb = $xl.Workbooks.Open($ExcelFilePath)
        }
        else {
        
            $xb = $xl.Workbooks.Add()
            $xb.SaveAs($ExcelFilePath)
       
        }

        $ws = $xb.Sheets.Item(1)
    }

    # Set counter Variables
    $Script:rowcounter = 1 
    $counter = 1

    # Set Variable for the top level group
    $search = get-adgroup -filter {name -like $TopLevelGroup}
    Write-Host "`n`n"
    
    # Set variable for the nested groups of the top level group
    $groups = get-adgroupmember $search.objectGUID | where {$_.objectClass -like "Group"}
    write-host -ForegroundColor White $TopLevelGroup
    
    # If outputting to excel, put top level group in cell 1,1
    if ($WithExcel) {
        $ws.cells.item($rowcounter, $counter) = $TopLevelGroup
    }

    $counter++
    Increment-Rowcounter

    # Recursively get nested groups
    $groups | ForEach-Object {Find-NestedGroups $counter $Script:rowcounter}

    # Release Excel COM Object if it was used.
    if ($WithExcel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
    }
}




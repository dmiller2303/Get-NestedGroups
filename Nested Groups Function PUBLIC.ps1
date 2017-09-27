function Get-NestedGroups {

    [CmdletBinding()]
    Param(
        
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [String]$TopLevelGroup,

        [Parameter(ParameterSetName='Excel',Mandatory=$false)]
        [Switch]$WithExcel,

        [parameter(ParameterSetName='Excel',Mandatory=$true)]
        [String]$ExcelFilePath

    )

    #Switch statement to put items into the Excel cells, color the cells, and output to the console
    function ExcelSwitchy ($counter, $rowcounter) {
        switch ($counter) {
            3 {

                Write-Host -ForegroundColor Green "      $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 4
                $Script:rowcounter++
                
            }

            4 {
                Write-Host -ForegroundColor Cyan "         $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 8
                $Script:rowcounter++
                
            }

            5 {
                Write-Host -ForegroundColor Magenta "            $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 18
                $Script:rowcounter++
            }

            6 {
                Write-Host -ForegroundColor Gray "               $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 16
                $Script:rowcounter++
            }

            7 {
                Write-Host -ForegroundColor red "                  $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 3
                $Script:rowcounter++
            }

            8 {
                Write-Host -ForegroundColor darkyellow "                     $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 45
                $Script:rowcounter++
            }

            9 {
                Write-Host -ForegroundColor DarkGreen "                        $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 10
                $Script:rowcounter++
            }

            10 {
                Write-Host -ForegroundColor Darkcyan "                           $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 33
                $Script:rowcounter++
            }

            11 {
                Write-Host -ForegroundColor DarkMagenta "                              $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 29
                $Script:rowcounter++
            }

            12 {
                Write-Host -ForegroundColor DarkGray "                                 $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 56
                $Script:rowcounter++
            }

            13 {
                Write-Host -ForegroundColor DarkRed "                                    $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 30
                $Script:rowcounter++
            }

            14 {
                Write-Host -ForegroundColor White "                                       $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 2
                $Script:rowcounter++
            }

            15 {
                Write-Host -ForegroundColor black "                                          $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 51
                $Script:rowcounter++
            }

            16 {
                Write-Host -ForegroundColor darkyellow "                                             $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 45
                $Script:rowcounter++
            }

            17-30 {
                
                Write-Host "need more swiches"
                $ws.cells.item($rowcounter, $counter) = $groupName
        
            }

            default {
                Write-Host -foregroundcolor Yellow "   $groupName"
                $ws.cells.item($rowcounter, $counter) = $groupName
                $ws.cells.item($rowcounter, $counter).interior.colorIndex = 6
                $Script:rowcounter++
            }
        }
    }

    #Switch Statment to output the groups to the console
    function Switchy ($counter){
    switch ($counter){
        3 {Write-Host -ForegroundColor Green "      $groupName"}
        4 {Write-Host -ForegroundColor Cyan "         $groupName"}
        5 {Write-Host -ForegroundColor Magenta "            $groupName"}
        6 {Write-Host -ForegroundColor Gray "               $groupName"}
        7 {Write-Host -ForegroundColor red "                  $groupName"}
        8 {Write-Host -ForegroundColor darkyellow "                     $groupName"}
        9 {Write-Host -ForegroundColor DarkGreen "                        $groupName"}
        10 {Write-Host -ForegroundColor Darkcyan "                           $groupName"}
        11 {Write-Host -ForegroundColor DarkMagenta "                              $groupName"}
        12 {Write-Host -ForegroundColor DarkGray "                                 $groupName"}
        13 {Write-Host -ForegroundColor DarkRed "                                    $groupName"}
        14 {Write-Host -ForegroundColor White "                                       $groupName"}
        15 {Write-Host -ForegroundColor black "                                          $groupName"}
        16 {Write-Host -ForegroundColor darkyellow "                                             $groupName"}
        17-30 {Write-Host "need more swiches"}
        default {Write-Host -foregroundcolor Yellow "   $groupName"}
      }
    }


    #Function that runs against each group in the pipeline that gathers its Nested Groups and then calls itself against the new Nested Groups in the pipeline until there are no nested groups
    function NestedGroups{
    
        $groupName = $_.name
        $objectGUID = $_.objectGUID
        $group = get-adgroupmember $_.objectGUID | where {$_.objectClass -like "Group"} 
        
        try {$h = get-adgroupmember $group | where {$_.objectClass -like "Group"}
        }catch {}

        if($WithExcel){

            ExcelSwitchy $counter $script:rowcounter
       
        }else{

            Switchy $counter
       
        }
    
        #Prevents an infinte loop of nested groups by checking to see if the 'child' group is also the 'parent' of the current group
        if ($objectGUID -ccontains $h.objectGUID ) {

            $counter++
        
            $groupName = (Get-ADGroupMember $objectGUID | where {$_.objectClass -like "Group"}).name

             if($WithExcel){
                  ExcelSwitchy $counter
             }else{
                  Switchy $counter
             }
        }
        else {
    
            $counter++
            $group | % {NestedGroups}
        
        }
    }


#End of Functions, Starting point.

    #If outputting to Excel, starts excel and opens or creates the file to be used.
    if($WithExcel){

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $true

        if(Test-Path $ExcelFilePath){
            $xb = $xl.Workbooks.Open($ExcelFilePath)
        }else{
        
            $xb = $xl.Workbooks.Add()
            $xb.saveas($ExcelFilePath)
       
        }

        $ws = $xb.Sheets.Item(1)
     }

    #Set counter Variables
    $Script:rowcounter = 1 
    $Script:counter = 1

    #Set Variable for the top level group
    $search = get-adgroup -filter {name -like $TopLevelGroup}
    Write-Host "`n`n"
    
    #Set variable for the nested groups of the top level group
    $groups = get-adgroupmember $search.objectGUID | where {$_.objectClass -like "Group"}
    write-host -ForegroundColor White $TopLevelGroup
    
    #If outputting to excel, put top level group in cell 1,1
    if($WithExcel){
        $ws.cells.item($Script:rowcounter, $counter) = $TopLevelGroup
    }

    $Script:counter++
    $Script:rowcounter++

    #Recursively get nested groups
    $groups | % {NestedGroups}

    #Release Excel COM Object if it was used.
    if($WithExcel){
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
    }
}




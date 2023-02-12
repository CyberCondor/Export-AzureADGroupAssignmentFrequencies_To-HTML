<#
.SYNOPSIS
Groups typically control permissions and sepperation of duties.
Information produced by this program can help with understanding the state of permission management.
To find out title and department demographics of each Azure AD Group with the help of frequency analysis.
.DESCRIPTION
Analyze Azure Active Directory Groups by frequency of title and department
.EXAMPLE
PS C:\> Export-AzureADGroupAssignmentFrequencies_To-HTML.ps1

## Input
Template.html
- Template for HTML layout 
styles.css
- Styles to display HTML layout properly
Users
- Contains list of all users
Groups
- Contains list of all groups and their members

## Output
~\_FrequenyAnalysisReports\AzureADGroupAssignmentFrequenyAnalysisReport_$Date
- Groups.html
- Users.html
- JobTitles.html
- Departments.html
- styles.css

#>
Write-Host "`n`t`tAttempting to query Azure Active Directory." -BackgroundColor Black -ForegroundColor Yellow
try{Get-AzureADUser -All $true > $null -ErrorAction stop
}
catch{$errMsg = $_.Exception.message
    if($errMsg.Contains("is not recognized as the name of a cmdlet")){
        Write-Warning "`t $_.Exception"
        Write-Output "Ensure 'AzureAD PS Module is installed. 'Install-Module AzureAD'"
        break
    }
    elseif($_.Exception -like "*Connect-AzureAD*"){
        Write-Warning "`t $_.Exception"
        Write-Output "Calling Connect-AzureAD"
        try{Connect-AzureAD -ErrorAction stop
        }
        catch{$errMsg = $_.Exception.message
            Write-Warning "`t $_.Exception"
            break
        }
    }
    else{Write-Warning "`t $_.Exception" ; break}
}

function Get-ExistingUsers_AzureAD{
    try{$ExistingUsers = Get-AzureADUser -All $true -ErrorAction Stop
        return $ExistingUsers
    }
    catch{$errMsg = $_.Exception.message
        Write-Warning "`t $_.Exception"
        return $null
    }
}

function Create-NewHtmlReportFromTemplate($TemplateFilePath, $NameOfPage, $ReportFolder, $divInfo, $NavLinkNames, $InfoTitles){
    if(!(Test-Path $TemplateFilePath)){return $null}
    if(!(Test-Path $ReportFolder.Path)){return $null}

    $NewHtmlReport = New-Object -TypeName PSObject -Property @{Name="$($NameOfPage).html";Path="$($ReportFolder.Path)\$($NameOfPage).html"}

    $InfoDivTitles = $(foreach($i in $InfoTitles){("<a href=#$($i)> $($i) </a>")})

    $H2AndDivInfo = "            <h2 class='glow'>
                $($NameofPage)
            </h2>
            <p class='div-info'>
                $(foreach($i in $divInfo){($i + '<br>')})
            </p>"

    Get-Content $TemplateFilePath | Select -First ((Select-String "<title>" $TemplateFilePath | select -First 1).linenumber - 1) | Out-File $NewHtmlReport.Path -Encoding utf8
    Write-Output "        <title>$($NameofPage)</title>" | Out-File $NewHtmlReport.Path -Append -Encoding utf8

    Get-Content $TemplateFilePath | Select -First ((Select-String "<h1>" $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "<h1>" $TemplateFilePath | select -First 1).linenumber - 1) - ((Select-String "<title>" $TemplateFilePath | select -First 1).linenumber)) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "                        <h1>$($ReportFolder.Name)</h1>" | Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "                        <p class='info-div-titles'>$($InfoDivTitles)</p>"       | Out-File $NewHtmlReport.Path -Append -Encoding utf8

    Get-Content $TemplateFilePath | Select -First ((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber) - ((Select-String "</p>" $TemplateFilePath | select -First 1).linenumber) - 1) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    foreach($NavLink in $NavLinkNames){
        if(!(Get-Content $TemplateFilePath | Select-String "href='$($NavLink).html'")){
            Write-Output "                            <li class='nav-link'><a href='$($NavLink).html' alt='$($NavLink)'>$($NavLink)</a></li>" | 
                Out-File $NewHtmlReport.Path -Append -Encoding utf8
        }
    }
    
    Get-Content $TemplateFilePath | Select -First ((Select-String "<h2 " $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "<h2 " $TemplateFilePath | select -First 1).linenumber) - ((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber)) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "$($H2AndDivInfo)" | Out-File $NewHtmlReport.Path -Append -Encoding utf8
    
    return $NewHtmlReport
}
function New-InfoDiv($HtmlReportPath, $Title, $Info){
    if(!(Test-Path $HtmlReportPath)){return $null}
    $Output = "            <div class='info-div'>
                <h3 id='$($Title)'>$($Title)</h3>
                <p class='div-info'>
                    $(foreach($i in $Info){($i + '<br>')})
                </p>"
    write-Output $Output | Out-File $HtmlReportPath -Append -Encoding utf8
}
function ExportTo-HTML_NewTable-Div($Title, $Info, $Object){
    $htmlParams = @{
      PreContent = "                <div class='table-div'>
                    <h4>$($Title)</h4>
                    <p class='table-info'>
                        Total: $($Info)
                    </p>"
      PostContent = "                </div>"
    }
   return $Object | ConvertTo-Html -As Table -Fragment @htmlParams
}
function ExportTo-HTML_NewTable-Div_Dynamic($Title, $Info, $Description, $Object){
    $htmlParams = @{
      PreContent = "                <div class='table-div'>
                    <h4>$($Title)</h4>
                    <p class='table-info'>
                        $($Description)
                        $($Info | ConvertTo-Html -As List -Fragment)
                    </p>"
      PostContent = "                </div>"
    }
   return $Object | ConvertTo-Html -As Table -Fragment @htmlParams
}
function End-InfoDiv($HtmlReportPath){
    if(!(Test-Path $HtmlReportPath)){return $null}
    Write-Output "            </div>" | Out-File $HtmlReportPath -Append -Encoding utf8
}
function Append-HtmlReport_Footer($HtmlReportPath){
    $Footer = "            <footer class='footnav'>
                Report Produced by $($env:USERNAME) on hostname: $(hostname)
                $(Get-Date -Format yyy-MM-dd)
                CONFIDENTIAL - NOT FOR DISTRIBUTION
            </footer>
        </div>
    </body>
</html>"
    $Footer  | Out-File $HtmlReportPath -Append -Encoding utf8
}

function Get-PropertyFrequencies($Property, $Object){
    $Total = ($Object).count
    $ProgressCount = 0
    $AllUniquePropertyValues = $Object | select $Property | sort $Property | unique -AsString # Get All Uniques
    $PropertyFrequencies = @()                                                                # Init empty Object
    $isDate = $false                                                                                                                                                          
    foreach($UniqueValue in $AllUniquePropertyValues){
        if(!($isDate -eq $true)){
            if([string]$UniqueValue.$Property -as [DateTime]){$isDate = $true}
        }
        $PropertyFrequencies += New-Object -TypeName PSobject -Property @{$Property=$($UniqueValue.$Property);Count=0;Frequency="100%"} # Copy Uniques to Object Array and Init Count as 0
    }
    if(($isDate -eq $true) -and (($Object | Select $Property | Get-Member).Definition -like "*datetime*")){
        foreach($PropertyFrequency in $PropertyFrequencies){
            if(($PropertyFrequency.$Property) -and ([string]$PropertyFrequency.$Property -as [DateTime])){
                try{$PropertyFrequency.$Property = $PropertyFrequency.$Property.ToString("yyyy-MM")}
                catch{# Nothing
                }
            }
        }
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
                else{
                    try{if($PropertyName.ToString("yyyy-MM") -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}}
                    catch{# Nothing
                    }
                }
            }
        }
    }
    else{
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
            }
        }
    }
    Write-Progress -id 1 -Completed -Activity "Complete"
    if($Total -gt 0){
        foreach($PropertyFrequency in $PropertyFrequencies){$PropertyFrequency.Frequency = ($PropertyFrequency.Count/$Total).ToString("P")}
    }
    return $PropertyFrequencies | select Count,$Property,Frequency | sort Count,$Property | Unique -AsString
}
function DisplayFrequencies($Property, $PropertyFrequencies){
    return $PropertyFrequencies | select Count,$Property,Frequency | sort Count,$Property
}

function Index-AzureADGroupAssignments{
    $Groups = Get-AzureADGroup -All $true
    $TotalGroups = ($Groups).count
    if($TotalGroups -gt 0){
        $ProgressCount = 1/$TotalGroups
        foreach($Group in $Groups){
            Write-Progress -id 2 -Activity "Indexing All Azure AD Group Assignments -> ( $([int]$ProgressCount) / $TotalGroups )" -Status "$(($ProgressCount++/$TotalGroups).ToString("P")) Complete"
            $Group | Add-Member -NotePropertyMembers @{GroupMembers=$(Get-AzureADGroupMember -ObjectId $Group.ObjectId -All $true)}
        }
        Write-Progress -id 2 -Completed -Activity "Complete"
        return $Groups
    }
    else{return $null}
}

function Correlate-TitlesAndDepartmentsInEachGroup($Groups, $ExistingUsers_AzureAD){
    foreach($Group in $Groups){
        $FoundGroupMembers = $false
        $JobTitlesInThisGroup = @()
        $DepartmentsInThisGroup = @()
        foreach($ExistingUser in $ExistingUsers_AzureAD){
            foreach($GroupMember in $Group.GroupMembers){
                if($GroupMember.UserPrincipalName -eq $ExistingUser.UserPrincipalName){
                    $FoundGroupMembers = $true
                    
                    if($DepartmentsInThisGroup){
                        $DepartmentFound = $false
                        foreach($DepartmentInThisGroup in $DepartmentsInThisGroup){
                            if($DepartmentInThisGroup.Department -ceq $ExistingUser.Department){
                                $DepartmentFound = $true
                                $DepartmentInThisGroup.Count++
                            }
                        }
                        if($DepartmentFound -eq $false){$DepartmentsInThisGroup += New-Object -TypeName PSobject -Property @{Count=1;Department=$ExistingUser.Department;Frequency="100%"}}
                    }
                    else{$DepartmentsInThisGroup += New-Object -TypeName PSobject -Property @{Count=1;Department=$ExistingUser.Department;Frequency="100%"}}
                    
                    if($JobTitlesInThisGroup){
                        $TitleFound = $false
                        foreach($JobTitleInThisGroup in $JobTitlesInThisGroup){
                            if($JobTitleInThisGroup.JobTitle -ceq $ExistingUser.JobTitle){
                                $TitleFound = $true
                                $JobTitleInThisGroup.Count++
                            }
                        }
                        if($TitleFound -eq $false){$JobTitlesInThisGroup += New-Object -TypeName PSobject -Property @{Count=1;JobTitle=$ExistingUser.JobTitle;Frequency="100%"}}
                    }
                    else{$JobTitlesInThisGroup += New-Object -TypeName PSobject -Property @{Count=1;JobTitle=$ExistingUser.JobTitle;Frequency="100%"}}
                }
            }
        }
        if(($FoundGroupMembers -eq $true) -and (($Group.GroupMembers.Department).count -eq 0)){Write-Warning "FoundGroupMembers = $FoundGroupMembers & Group.GroupMembers.department count = $(($Group.GroupMembers.Department).count) - $($Group.DisplayName)"}
        
        if(($FoundGroupMembers -eq $false) -or (($Group.GroupMembers.Department).count -eq 0)){
            $Group | Add-Member -NotePropertyMembers @{Departments="NULL"}
            $Group | Add-Member -NotePropertyMembers @{JobTitles="NULL"}
        }
        elseif(($FoundGroupMembers -eq $true) -and (($Group.GroupMembers.Department).count -gt 0)){
            foreach($DepartmentInThisGroup in $DepartmentsInThisGroup){
                $DepartmentInThisGroup.Frequency = ($DepartmentInThisGroup.Count/($Group.GroupMembers.Department).count).ToString("P")
            }
            foreach($JobTitleInThisGroup in $JobTitlesInThisGroup){
                $JobTitleInThisGroup.Frequency = ($JobTitleInThisGroup.Count/($Group.GroupMembers.JobTitle).count).ToString("P")
            }
            $Group | Add-Member -NotePropertyMembers @{Departments=$DepartmentsInThisGroup}
            $Group | Add-Member -NotePropertyMembers @{JobTitles=$JobTitlesInThisGroup}
        }
    }
}
function Correlate-GroupsInEachTitleAndDepartment($ExistingUsers_AzureAD, $JobTitles, $Departments){
    foreach($JobTitle in $JobTitles){
        $FoundGroupsInTitle = $false
        $GroupsInThisJobTitle = @()
        foreach($ExistingUser in $ExistingUsers_AzureAD | where{$_.JobTitle -ceq $JobTitle.JobTitle}){
            foreach($Group in $ExistingUser.AssignedAzureADGroups){
                $FoundGroupsInTitle = $true
                if($GroupsInThisJobTitle){
                    $GroupFound = $false
                    foreach($GroupInThisJobTitle in $GroupsInThisJobTitle){
                        if($GroupInThisJobTitle.GroupName -ceq $Group.DisplayName){
                            $GroupFound = $true
                            $GroupInThisJobTitle.Count++
                        }
                    }
                    if($GroupFound -eq $false){$GroupsInThisJobTitle += New-Object -TypeName PSobject -Property @{Count=1;GroupName=$Group.DisplayName;Frequency="100%"}}
                }
                else{$GroupsInThisJobTitle += New-Object -TypeName PSobject -Property @{Count=1;GroupName=$Group.DisplayName;Frequency="100%"}}  
            }
        }
        if($FoundGroupsInTitle -eq $false){
            $JobTitle | Add-Member -NotePropertyMembers @{Groups="NULL"}
        }
        elseif($FoundGroupsInTitle -eq $true){
            foreach($GroupInThisJobTitle in $GroupsInThisJobTitle){
                $GroupInThisJobTitle.Frequency = ($GroupInThisJobTitle.Count/$JobTitle.count).ToString("P")
            }
            $JobTitle | Add-Member -NotePropertyMembers @{Groups=$GroupsInThisJobTitle}
        }
    }  
    foreach($Department in $Departments){
        $FoundGroupsInDepartment = $false
        $GroupsInThisDepartment = @()
        foreach($ExistingUser in $ExistingUsers_AzureAD | where{$_.Department -ceq $Department.Department}){
            foreach($Group in $ExistingUser.AssignedAzureADGroups){
                $FoundGroupsInDepartment = $true
                if($GroupsInThisDepartment){
                    $GroupFound = $false
                    foreach($GroupInThisDepartment in $GroupsInThisDepartment){
                        if($GroupInThisDepartment.GroupName -ceq $Group.DisplayName){
                            $GroupFound = $true
                            $GroupInThisDepartment.Count++
                        }
                    }
                    if($GroupFound -eq $false){$GroupsInThisDepartment += New-Object -TypeName PSobject -Property @{Count=1;GroupName=$Group.DisplayName;Frequency="100%"}}
                }
                else{$GroupsInThisDepartment += New-Object -TypeName PSobject -Property @{Count=1;GroupName=$Group.DisplayName;Frequency="100%"}}   
            }   
        }
        if($FoundGroupsInDepartment -eq $false){
            $Department | Add-Member -NotePropertyMembers @{Groups="NULL"}
        }
        elseif($FoundGroupsInDepartment -eq $true){
            foreach($GroupInThisDepartment in $GroupsInThisDepartment){
                $GroupInThisDepartment.Frequency = ($GroupInThisDepartment.Count/$Department.count).ToString("P")
            }
            $Department | Add-Member -NotePropertyMembers @{Groups=$GroupsInThisDepartment}
        }  
    } 
}

function Add-AzureADGroupsAssigned($Groups, $ExistingUser){
    $AssignedGroups = @()
    foreach($Group in $Groups){
        foreach($GroupMember in $Group.GroupMembers){
            if($GroupMember.UserPrincipalName -eq $ExistingUser.UserPrincipalName){
                $AssignedGroups += $Group | Select DisplayName,MailEnabled,SecurityEnabled,MailNickName,Mail,LastDirSyncTime,Description
                break #If user found in this group, just move on to search the next Group
            }
        }
    }
    $ExistingUser | Add-Member -NotePropertyMembers @{AssignedAzureADGroups=$AssignedGroups}
}

function Get-FrequencyObjectInEachGroup($Groups, $GroupAssignmentFrequencies, $FrequencyObjectName, $Property, $HtmlReportPath){
    foreach($Group in $Groups | sort GroupMembers,Description){
        if($Group.$FrequencyObjectName -ne "NULL"){
            foreach($GroupFrequency in $GroupAssignmentFrequencies){
                if($Group.DisplayName -ceq $GroupFrequency.DisplayName){
                    ExportTo-HTML_NewTable-Div_Dynamic "$(($Group.$FrequencyObjectName).Count) $FrequencyObjectName in: $($Group.DisplayName)" $($GroupFrequency | Select DisplayName,Frequency,Count | sort Count,DisplayName) $($Group.Description) $($Group.$FrequencyObjectName | Select Count,$Property,Frequency | sort Count,$Property) |
                        Out-File $HtmlReportPath -Append -Encoding utf8
                    #Write-Output "$(($Group.$FrequencyObjectName).Count) $FrequencyObjectName in: $($Group.DisplayName)"
                    #$GroupFrequency | Select DisplayName,Frequency,Count | sort Count,DisplayName |fl
                    #if($Group.Description){Write-Output "$($Group.Description)"}
                    #$Group.$FrequencyObjectName | Select Count,$Property,Frequency | sort Count,$Property | ft
                }
            }
        }
    }
}
function Get-GroupsInEachPropertyFrequencyObject($PropertyFrequencyObject, $Property, $HtmlReportPath){
    foreach($Thing in $PropertyFrequencyObject | sort Groups,Count,$Property){
        if($Thing.Groups -ne "NULL"){
            ExportTo-HTML_NewTable-Div_Dynamic "$($Thing.Groups.Count) Groups in: $($Thing.$Property)" $($Thing | Select $Property,Frequency,Count) $null $($Thing.Groups | Select Count,GroupName,Frequency | sort Count,GroupName) |
                Out-File $HtmlReportPath -Append -Encoding utf8
            #Write-Output "$($Thing.Groups.Count) Groups in: $($Thing.$Property)"
            #$Thing | Select $Property,Frequency,Count | fl
            #$Thing.Groups | Select Count,GroupName,Frequency | sort Count,GroupName | ft
        }
    }
}

function main{
    $ExistingUsers_AzureAD = Get-ExistingUsers_AzureAD #| where {$_.AccountEnabled -eq $true}
    if($ExistingUsers_AzureAD -eq $null){break}

    $MainFolderPath   = "~\_FrequenyAnalysisReports"
    $ReportFolderName = ("AzureADGroupAssignmentFrequenyAnalysisReport_" + $(get-date -format yyy-MM-dd))
    $ReportFolder     = New-Object -TypeName PSObject -Property @{Name=$ReportFolderName;Path="$MainFolderPath\$ReportFolderName"}
    $CurrDir          = (pwd).Path
    if((!(Test-Path Template.html)) -or (!((Get-Content Template.html | select -First 1).Contains("<!DOCTYPE HTML>")))){
        Write-Output "Cannot Find HTML Template File."
        break
    }
    if(!(Test-Path $MainFolderPath)){
        mkdir $MainFolderPath
    }
    if(!(Test-Path $ReportFolder.Path)){
        mkdir $ReportFolder.Path
        try{Copy "$CurrDir/styles.css" $ReportFolder.Path
        }
        catch{$errMsg = $_.Exception.message
            Write-Warning "`t $_.Exception"
            break
        }
    }
    else{
        Write-Warning "Folder '$($ReportFolder.Path)' already exists.`nProgram already ran today."
        break
    }

    Write-Output "`n$($env:UserName) - Started this program @ $(date)"
    Write-Host "$($env:UserName) - Started this program @ $(date)"

    $Departments = Get-PropertyFrequencies "Department" $ExistingUsers_AzureAD
    $JobTitles   = Get-PropertyFrequencies "JobTitle" $ExistingUsers_AzureAD
    
    $Groups = Index-AzureADGroupAssignments
    if($Groups -eq $null){Write-Output "Groups is null";break}

    foreach($ExistingUser in $ExistingUsers_AzureAD){Add-AzureADGroupsAssigned $Groups $ExistingUser}

    Correlate-TitlesAndDepartmentsInEachGroup $Groups $ExistingUsers_AzureAD
    Correlate-GroupsInEachTitleAndDepartment $ExistingUsers_AzureAD $JobTitles $Departments
    
    ################################################################
        #Sub-Objects
    #---
    #Groups
    $GroupAssignmentFrequencies        = Get-PropertyFrequencies "DisplayName" $ExistingUsers_AzureAD.AssignedAzureADGroups
        $GroupDisplayNameFrequencies   = Get-PropertyFrequencies "DisplayName" $Groups
    $DuplicateGroupNameFrequencies     = DisplayFrequencies "DisplayName" ($GroupDisplayNameFrequencies | where{$_.Count -gt 1})
    $AmbiguousGroupNameAssignments = @()
    foreach($GroupAssignment in $GroupAssignmentFrequencies){
        foreach($Group in $GroupDisplayNameFrequencies | where{$_.Count -gt 1}){
            if($GroupAssignment.DisplayName -eq $Group.DisplayName){
                $AmbiguousGroupNameAssignments += $GroupAssignment
            }
        }
    }
        #DepartmentFrequenciesInEachGroup**
    #** Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "Departments" "Department"
        #JobTitleFrequenciesInEachGroup**
    #** Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "JobTitles" "JobTitle"
    $GroupsWithoutMembers              = $Groups | where{$_.GroupMembers -eq $null} | 
        select DisplayName,MailEnabled,SecurityEnabled,MailNickName,Mail,LastDirSyncTime,Description | sort MailEnabled,SecurityEnabled,Description,DisplayName
    $GroupsWithoutUsers_ButWithMembers = $Groups | where{($_.Departments -eq "NULL") -and ($_.GroupMembers -ne $null)} | 
        select DisplayName,MailEnabled,SecurityEnabled,MailNickName,Mail,LastDirSyncTime,Description | sort MailEnabled,SecurityEnabled,Description,DisplayName
    #---
    #Users
    $EnabledUserFrequencies    = Get-PropertyFrequencies "AccountEnabled" $ExistingUsers_AzureAD
    $DirSyncEnabledFrequencies = Get-PropertyFrequencies "DirSyncEnabled" $ExistingUsers_AzureAD
    $CountryFrequencies        = Get-PropertyFrequencies "Country" $ExistingUsers_AzureAD
    $StateFrequencies          = Get-PropertyFrequencies "State" $ExistingUsers_AzureAD
    $OfficeFrequencies         = Get-PropertyFrequencies "Office" $ExistingUsers_AzureAD

    $EnabledUsersWithoutGroupAssignments = @()
    foreach($ExistingUser in $ExistingUsers_AzureAD | where{$_.AccountEnabled -eq $true}){
        if(!($ExistingUser.AssignedAzureADGroups)){$EnabledUsersWithoutGroupAssignments += $ExistingUser}
    }
    $DisabledUsersWithGroupAssignments   = @()
    foreach($ExistingUser in $ExistingUsers_AzureAD | where{$_.AccountEnabled -eq $false}){
        if($ExistingUser.AssignedAzureADGroups){$DisabledUsersWithGroupAssignments      += $ExistingUser}
    }
    #---
    #JobTitles
	    #JobTitleFrequencies              - DisplayFrequencies "JobTitle" $JobTitles
    #$JobTitles                       = Get-PropertyFrequencies "JobTitle" $ExistingUsers_AzureAD
        #JobTitlesWithoutGroupAssignments - DisplayFrequencies "JobTitle" $JobTitlesWithoutGroupAssignments
    $JobTitlesWithoutGroupAssignments = $JobTitles | where{$_.Groups -eq "NULL"}
	    #GroupAssignmentFrequenciesInEachJobTitle**
    #** Get-GroupsInEachPropertyFrequencyObject $JobTitles "JobTitle"
    #---
    #Departments
	    #DepartmentFrequencies              - DisplayFrequencies "Department" $Departments
    #$Departments                       = Get-PropertyFrequencies "Department" $ExistingUsers_AzureAD
	    #DepartmentsWithoutGroupAssignments - DisplayFrequencies "Department" $DepartmentsWithoutGroupAssignments
    $DepartmentsWithoutGroupAssignments = $Departments | where{$_.Groups -eq "NULL"}
	    #GroupAssignmentFrequenciesInEachDepartment**
    #** Get-GroupsInEachPropertyFrequencyObject $Departments "Department"
    #---
    #End Sub-Objects
    ###
    ################################################################
        #Totals
    #---
    #Groups
    $TotalGroups                              = ($Groups).Count
    $TotalUniqueGroupNames                    = ($Groups | select DisplayName | sort DisplayName | unique -AsString).count
	    #TotalDuplicateGroupNames
        $TotalDuplicateGroupNamesTotalCount = 0
        foreach($Group in $GroupDisplayNameFrequencies | where{$_.Count -gt 1}){$TotalDuplicateGroupNamesTotalCount += $Group.Count}
                    # TotalGroups = (TotalUniqueGroupNames - (GroupDisplayNameFrequencies | where{$_.Count -gt 1}).count) + TotalDuplicateGroupNames
    $TotalDuplicateGroupNames                 = ($TotalDuplicateGroupNamesTotalCount - ($GroupDisplayNameFrequencies | where{$_.Count -gt 1}).count)
                    # TotalGroups = TotalUniqueGroupNames + TotalDuplicateGroupNames
    $TotalMailEnabledGroups                   = ($Groups | where{$_.MailEnabled -eq $true}).count         
    $TotalSecurityEnabledGroups               = ($Groups | where{$_.SecurityEnabled -eq $true}).count
    $TotalMailAndSecurityEnabledGroups        = ($Groups | where{($_.MailEnabled -eq $true) -and ($_.SecurityEnabled -eq $true)}).count        
    $TotalGroupsWithADescription              = ($Groups | where{$_.Description -ne $null}).count    
    $TotalGroupsWithoutADescription           = ($Groups | where{$_.Description -eq $null}).count      
    $TotalGroupsWithDirSyncEnabled            = ($Groups | where{$_.DirSyncEnabled -eq $true}).count
    $TotalGroupsWithoutDirSyncEnabled         = ($Groups | where{$_.DirSyncEnabled -ne $true}).count           
    $TotalGroupsWithMembers                   = ($Groups | where{$_.GroupMembers -ne $null}).count         
    $TotalGroupsWithoutMembers                = ($GroupsWithoutMembers).count
    $TotalGroupsWithUsers                     = ($Groups | where{$_.Departments -ne "NULL"}).count
    $TotalGroupsWithoutUsers                  = ($Groups | where{$_.Departments -eq "NULL"}).count
    $TotalGroupsWithoutUsers_ButWithMembers   = ($GroupsWithoutUsers_ButWithMembers).count
    $TotalGroupsWhereLastDirSyncTimeGT180Days = ($Groups | where{$_.LastDirSyncTime -lt [datetime]::Now.AddDays(-180)} | select DisplayName,LastDirSyncTime | sort LastDirSyncTime).count
    #---
    #Users
    $TotalUsers                                               = ($ExistingUsers_AzureAD).count
    $TotalEnabledUsers                                        = ($ExistingUsers_AzureAD| where{$_.AccountEnabled -eq $true}).count
    $TotalDisabledUsers                                       = ($ExistingUsers_AzureAD| where{$_.AccountEnabled -eq $false}).count
	    #TotalUsersWithGroupAssignments
        $TotalUsersWithGroupAssignments = 0
        $TotalUsersWithoutGroupAssignments = 0
        foreach($ExistingUser in $ExistingUsers_AzureAD){
            if($ExistingUser.AssignedAzureADGroups){$TotalUsersWithGroupAssignments++}
            else{$TotalUsersWithoutGroupAssignments++}
        }
    #$TotalUsersWithGroupAssignments
	    #TotalUsersWithoutGroupAssignments ^
    #$TotalUsersWithoutGroupAssignments

    $TotalDisabledUsersWithGroupAssignments                   = ($DisabledUsersWithGroupAssignments).count
    $TotalEnabledUsersWithoutGroupAssignments                 = ($EnabledUsersWithoutGroupAssignments).count
    $TotalUsersWhereLastDirSyncTimeGT180DaysAndAccountEnabled = ($ExistingUsers_AzureAD | where{($_.LastDirSyncTime -lt [datetime]::Now.AddDays(-180)) -and ($_.AccountEnabled -eq $true) }).count
    #---
    #JobTitles
    $TotalUniqueJobTitles                  = ($JobTitles).Count
    $TotalJobTitlesWithGroupAssignments    = ($JobTitles.count - ($JobTitlesWithoutGroupAssignments).count)
    $TotalJobTitlesWithoutGroupAssignments = ($JobTitlesWithoutGroupAssignments).count
    #---
    #Departments
    $TotalUniqueDepartments                  = ($Departments).Count
    $TotalDepartmentsWithGroupAssignments    = ($Departments.count - ($DepartmentsWithoutGroupAssignments).count)
    $TotalDepartmentsWithoutGroupAssignments = ($DepartmentsWithoutGroupAssignments).count
    #---
    #End Totals
    ###
    ########################################
    #Make Sub-Info Before Displaying Layout
    $Groups_SubInfo = @("TotalGroups                             : $TotalGroups                             ",
	                    "TotalUniqueGroupNames                   : $TotalUniqueGroupNames                   ",
	                    "TotalDuplicateGroupNames                : $TotalDuplicateGroupNames                ",
	                    "TotalMailEnabledGroups                  : $TotalMailEnabledGroups                  ",
	                    "TotalSecurityEnabledGroups              : $TotalSecurityEnabledGroups              ",
	                    "TotalMailAndSecurityEnabledGroups       : $TotalMailAndSecurityEnabledGroups       ",
	                    "TotalGroupsWithADescription             : $TotalGroupsWithADescription             ",
	                    "TotalGroupsWithoutADescription          : $TotalGroupsWithoutADescription          ",
	                    "TotalGroupsWithDirSyncEnabled           : $TotalGroupsWithDirSyncEnabled           ",
	                    "TotalGroupsWithoutDirSyncEnabled        : $TotalGroupsWithoutDirSyncEnabled        ",
	                    "TotalGroupsWithMembers                  : $TotalGroupsWithMembers                  ",
	                    "TotalGroupsWithoutMembers               : $TotalGroupsWithoutMembers               ",
	                    "TotalGroupsWithUsers                    : $TotalGroupsWithUsers                    ",
	                    "TotalGroupsWithoutUsers                 : $TotalGroupsWithoutUsers                 ",
	                    "TotalGroupsWithoutUsers_ButWithMembers  : $TotalGroupsWithoutUsers_ButWithMembers  ",
	                    "TotalGroupsWhereLastDirSyncTimeGT180Days: $TotalGroupsWhereLastDirSyncTimeGT180Days")
    $Groups_SubInfo_Div1 = @("TotalGroups                           : $TotalGroups                           ",
		                     "TotalUniqueGroupNames                 : $TotalUniqueGroupNames                 ",
		                     "TotalDuplicateGroupNames              : $TotalDuplicateGroupNames              ",
		                     "TotalGroupsWithMembers                : $TotalGroupsWithMembers                ",
		                     "TotalGroupsWithoutMembers             : $TotalGroupsWithoutMembers             ",
		                     "TotalGroupsWithUsers                  : $TotalGroupsWithUsers                  ",
		                     "TotalGroupsWithoutUsers               : $TotalGroupsWithoutUsers               ",
		                     "TotalGroupsWithoutUsers_ButWithMembers: $TotalGroupsWithoutUsers_ButWithMembers")
    $Groups_SubInfo_Div2 = @("TotalUsers                              : $TotalUsers                              ",
		                     "TotalUsersWithGroupAssignments          : $TotalUsersWithGroupAssignments          ",
		                     "TotalUsersWithoutGroupAssignments       : $TotalUsersWithoutGroupAssignments       ",
		                     "TotalDisabledUsersWithGroupAssignments  : $TotalDisabledUsersWithGroupAssignments  ",
		                     "TotalEnabledUsersWithoutGroupAssignments: $TotalEnabledUsersWithoutGroupAssignments",
		                     "TotalGroupsWithUsers                    : $TotalGroupsWithUsers                    ",
		                     "TotalUniqueDepartments                  : $TotalUniqueDepartments                  ",
		                     "TotalDepartmentsWithGroupAssignments    : $TotalDepartmentsWithGroupAssignments    ",
		                     "TotalDepartmentsWithoutGroupAssignments : $TotalDepartmentsWithoutGroupAssignments ")
    $Groups_SubInfo_Div3 = @("TotalUsers                              : $TotalUsers                              ",
		                     "TotalUsersWithGroupAssignments          : $TotalUsersWithGroupAssignments          ",
		                     "TotalUsersWithoutGroupAssignments       : $TotalUsersWithoutGroupAssignments       ",
		                     "TotalDisabledUsersWithGroupAssignments  : $TotalDisabledUsersWithGroupAssignments  ",
		                     "TotalEnabledUsersWithoutGroupAssignments: $TotalEnabledUsersWithoutGroupAssignments",
		                     "TotalGroupsWithUsers                    : $TotalGroupsWithUsers                    ",
		                     "TotalUniqueJobTitles                    : $TotalUniqueJobTitles                    ",
		                     "TotalJobTitlesWithGroupAssignments      : $TotalJobTitlesWithGroupAssignments      ",
		                     "TotalJobTitlesWithoutGroupAssignments   : $TotalJobTitlesWithoutGroupAssignments   ")
    $Groups_SubInfo_Div4 = @("TotalGroupsWithMembers                : $TotalGroupsWithMembers                ",
		                     "TotalGroupsWithoutMembers             : $TotalGroupsWithoutMembers             ",
		                     "TotalGroupsWithUsers                  : $TotalGroupsWithUsers                  ",
		                     "TotalGroupsWithoutUsers               : $TotalGroupsWithoutUsers               ",
		                     "TotalGroupsWithoutUsers_ButWithMembers: $TotalGroupsWithoutUsers_ButWithMembers")
                                                                                                             
    $Users_SubInfo = @("TotalUsers                                              : $TotalUsers                                              ",
	                   "TotalEnabledUsers                                       : $TotalEnabledUsers                                       ",
	                   "TotalDisabledUsers                                      : $TotalDisabledUsers                                      ",
	                   "TotalUsersWithGroupAssignments                          : $TotalUsersWithGroupAssignments                          ",
	                   "TotalUsersWithoutGroupAssignments                       : $TotalUsersWithoutGroupAssignments                       ",
	                   "TotalDisabledUsersWithGroupAssignments                  : $TotalDisabledUsersWithGroupAssignments                  ",
	                   "TotalEnabledUsersWithoutGroupAssignments                : $TotalEnabledUsersWithoutGroupAssignments                ",
	                   "TotalUsersWhereLastDirSyncTimeGT180DaysAndAccountEnabled: $TotalUsersWhereLastDirSyncTimeGT180DaysAndAccountEnabled")
    $Users_SubInfo_Div1 = @("TotalEnabledUsers : $TotalEnabledUsers ",
		                    "TotalDisabledUsers: $TotalDisabledUsers")
    $Users_SubInfo_Div2 = @("TotalUsersWithGroupAssignments          : $TotalUsersWithGroupAssignments          ",
		                    "TotalUsersWithoutGroupAssignments       : $TotalUsersWithoutGroupAssignments       ",
		                    "TotalDisabledUsersWithGroupAssignments  : $TotalDisabledUsersWithGroupAssignments  ",
		                    "TotalEnabledUsersWithoutGroupAssignments: $TotalEnabledUsersWithoutGroupAssignments")

    $JobTitles_SubInfo = @("TotalUsers                           : $TotalUsers                           ",
	                       "TotalEnabledUsers                    : $TotalEnabledUsers                    ",
	                       "TotalDisabledUsers                   : $TotalDisabledUsers                   ",
	                       "TotalUniqueJobTitles                 : $TotalUniqueJobTitles                 ",
	                       "TotalJobTitlesWithGroupAssignments   : $TotalJobTitlesWithGroupAssignments   ",
	                       "TotalJobTitlesWithoutGroupAssignments: $TotalJobTitlesWithoutGroupAssignments")
    $JobTitles_SubInfo_Div1 = @("TotalUniqueJobTitles                 : $TotalUniqueJobTitles                 ",
		                        "TotalJobTitlesWithoutGroupAssignments: $TotalJobTitlesWithoutGroupAssignments")
    $JobTitles_SubInfo_Div2 = @("TotalJobTitlesWithGroupAssignments: $TotalJobTitlesWithGroupAssignments")

    $Departments_SubInfo = @("TotalUsers                             : $TotalUsers                             ",
	                         "TotalEnabledUsers                      : $TotalEnabledUsers                      ",
	                         "TotalDisabledUsers                     : $TotalDisabledUsers                     ",
	                         "TotalUniqueDepartments                 : $TotalUniqueDepartments                 ",
	                         "TotalDepartmentsWithGroupAssignments   : $TotalDepartmentsWithGroupAssignments   ",
	                         "TotalDepartmentsWithoutGroupAssignments: $TotalDepartmentsWithoutGroupAssignments")
    $Departments_SubInfo_Div1 = @("TotalUniqueDepartments                 : $TotalUniqueDepartments                 ",
		                          "TotalDepartmentsWithoutGroupAssignments: $TotalDepartmentsWithoutGroupAssignments")
    $Departments_SubInfo_Div2 = @("TotalDepartmentsWithGroupAssignments: $TotalDepartmentsWithGroupAssignments")
    #End Sub-Info
    ########################################
    #Make Reports from Template.html


    $Groups_InfoTitles      = @("GroupAssignmentsOverview","DepartmentFrequenciesInEachGroup","JobTitleFrequenciesInEachGroup","PossibleStaleGroups")
    $Users_InfoTitles       = @("GeneralInfo","PossiblyNeedsAttention")
    $JobTitles_InfoTitles   = @("JobTitlesOverview","GroupAssignmentFrequenciesInEachJobTitle")
    $Departments_InfoTitles = @("DepartmentsOverview","GroupAssignmentFrequenciesInEachDepartment")


    $MainObjectNames = @("Groups","Users","JobTitles","Departments")
    $Groups_HTML      = Create-NewHtmlReportFromTemplate "Template.html"      "Groups"      $ReportFolder $Groups_SubInfo      $MainObjectNames $Groups_InfoTitles     
    $Users_HTML       = Create-NewHtmlReportFromTemplate $Groups_HTML.Path    "Users"       $ReportFolder $Users_SubInfo       $MainObjectNames $Users_InfoTitles      
    $JobTitles_HTML   = Create-NewHtmlReportFromTemplate $Users_HTML.Path     "JobTitles"   $ReportFolder $JobTitles_SubInfo   $MainObjectNames $JobTitles_InfoTitles  
    $Departments_HTML = Create-NewHtmlReportFromTemplate $JobTitles_HTML.Path "Departments" $ReportFolder $Departments_SubInfo $MainObjectNames $Departments_InfoTitles

    #####################################################################
    #   LAYOUT
    # Export all info to HTML 
    ################################################## #Groups - Info-Div_Count: 4
    # Title: Groups
    # MainHeaderString: $FolderName
    # SubTitle: Groups
    # SubInfo: Groups_SubInfo
    # **Info-Div_1**
    #	- SubTitle: Group Assignments Overview
    #	- SubInfo: Groups_SubInfo_Div1
    New-InfoDiv              $Groups_HTML.Path "GroupAssignmentsOverview" $Groups_SubInfo_Div1
    #	- *Content:*
    #		- #GroupAssignmentFrequencies
    ExportTo-HTML_NewTable-Div "GroupAssignmentFrequencies" $($GroupAssignmentFrequencies).Count $(DisplayFrequencies "DisplayName" $GroupAssignmentFrequencies) | Out-File $Groups_HTML.Path -Append -Encoding utf8
    #		- #DuplicateGroupNameFrequencies
    ExportTo-HTML_NewTable-Div "DuplicateGroupNameFrequencies" $TotalDuplicateGroupNames $(DisplayFrequencies "DisplayName" $DuplicateGroupNameFrequencies) | Out-File $Groups_HTML.Path -Append -Encoding utf8
    #		- #AmbiguousGroupNameAssignments
    ExportTo-HTML_NewTable-Div "AmbiguousGroupNameAssignments" $($AmbiguousGroupNameAssignments).Count $(DisplayFrequencies "DisplayName" $AmbiguousGroupNameAssignments) | Out-File $Groups_HTML.Path -Append -Encoding utf8
    End-InfoDiv              $Groups_HTML.Path
    # **Info-Div_2**
    #	- SubTitle: DepartmentFrequenciesInEachGroup
    #	- SubInfo: Groups_SubInfo_Div2
    New-InfoDiv              $Groups_HTML.Path "DepartmentFrequenciesInEachGroup" $Groups_SubInfo_Div2
    #	- *Content:*
    #		- #DepartmentFrequenciesInEachGroup**
    Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "Departments" "Department" $Groups_HTML.Path
    End-InfoDiv              $Groups_HTML.Path
    # **Info-Div_3**
    #	- SubTitle: JobTitleFrequenciesInEachGroup
    #	- SubInfo: Groups_SubInfo_Div3
    New-InfoDiv              $Groups_HTML.Path "JobTitleFrequenciesInEachGroup" $Groups_SubInfo_Div3
    #	- *Content:*
    #		- #JobTitleFrequenciesInEachGroup**
    Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "JobTitles" "JobTitle" $Groups_HTML.Path
    End-InfoDiv              $Groups_HTML.Path
    # **Info-Div_4**
    #	- SubTitle: Possible Stale Groups
    #	- SubInfo: Groups_SubInfo_Div4
    New-InfoDiv              $Groups_HTML.Path "PossibleStaleGroups" $Groups_SubInfo_Div4
    #	- *Content:*
    #		- #GroupsWithoutMembers
    ExportTo-HTML_NewTable-Div "GroupsWithoutMembers" $TotalGroupsWithoutMembers $GroupsWithoutMembers | Out-File $Groups_HTML.Path -Append -Encoding utf8
    #		- #GroupsWithoutUsers_ButWithMembers
    ExportTo-HTML_NewTable-Div "GroupsWithoutUsers_ButWithMembers" $TotalGroupsWithoutUsers_ButWithMembers $GroupsWithoutUsers_ButWithMembers | Out-File $Groups_HTML.Path -Append -Encoding utf8
    End-InfoDiv              $Groups_HTML.Path
    #
    Append-HtmlReport_Footer $Groups_HTML.Path 
    ################################################## #Users - Info-Div_Count: 2
    # Title: Users
    # MainHeaderString: $FolderName
    # SubTitle: Users
    # SubInfo: Users_SubInfo
    # **Info-Div_1**
    #	- SubTitle: General Info
    #	- SubInfo: Users_SubInfo_Div1
    New-InfoDiv              $Users_HTML.Path "GeneralInfo" $Users_SubInfo_Div1
    #	- *Content:*
    #		- #EnabledUserFrequencies
    ExportTo-HTML_NewTable-Div "EnabledUserFrequencies" $null $(DisplayFrequencies "AccountEnabled" $EnabledUserFrequencies) | Out-File $Users_HTML.Path -Append -Encoding utf8
    #		- #DirSyncEnabledFrequencies
    ExportTo-HTML_NewTable-Div "DirSyncEnabledFrequencies" $null $(DisplayFrequencies "DirSyncEnabled" $DirSyncEnabledFrequencies) | Out-File $Users_HTML.Path -Append -Encoding utf8
    #		- #CountryFrequencies
    ExportTo-HTML_NewTable-Div "CountryFrequencies" ($CountryFrequencies).Count $(DisplayFrequencies "Country" $CountryFrequencies) | Out-File $Users_HTML.Path -Append -Encoding utf8
    #		- #StateFrequencies
    ExportTo-HTML_NewTable-Div "StateFrequencies" ($StateFrequencies).count $(DisplayFrequencies "State" $StateFrequencies) | Out-File $Users_HTML.Path -Append -Encoding utf8
    #		- #OfficeFrequencies
    ExportTo-HTML_NewTable-Div "OfficeFrequencies" ($OfficeFrequencies).count $(DisplayFrequencies "Office" $OfficeFrequencies) | Out-File $Users_HTML.Path -Append -Encoding utf8

    End-InfoDiv              $Users_HTML.Path
    # **Info-Div_2**
    #	- SubTitle: Possibly Needs Attention
    #	- SubInfo: Users_SubInfo_Div2
    New-InfoDiv              $Users_HTML.Path "PossiblyNeedsAttention" $Users_SubInfo_Div2
    #	- *Content:*
    #		- #EnabledUsersWithoutGroupAssignments
    ExportTo-HTML_NewTable-Div "EnabledUsersWithoutGroupAssignments" $TotalEnabledUsersWithoutGroupAssignments $($EnabledUsersWithoutGroupAssignments | select DisplayName,UserPrincipalName,UserType,Title,Department,LastDirSyncTime | sort UserType,DisplayName) | Out-File $Users_HTML.Path -Append -Encoding utf8
    #		- #DisabledUsersWithGroupAssignments
    ExportTo-HTML_NewTable-Div "DisabledUsersWithGroupAssignments" $TotalDisabledUsersWithGroupAssignments $($DisabledUsersWithGroupAssignments | select DisplayName,UserPrincipalName,UserType,Title,Department,LastDirSyncTime | sort UserType,DisplayName) | Out-File $Users_HTML.Path -Append -Encoding utf8
    End-InfoDiv              $Users_HTML.Path

    Append-HtmlReport_Footer $Users_HTML.Path
    ################################################## #JobTitles - Info-Div_Count: 2
    # Title: JobTitles
    # MainHeaderString: $FolderName
    # SubTitle: JobTitles
    # SubInfo: JobTitles_SubInfo
    # **Info-Div_1**
    #	- SubTitle: JobTitles Overview
    #	- SubInfo: JobTitles_SubInfo_Div1
    New-InfoDiv              $JobTitles_HTML.Path "JobTitlesOverview" $JobTitles_SubInfo_Div1
    #	- *Content:*
    #		- #JobTitleFrequencies
    ExportTo-HTML_NewTable-Div "JobTitleFrequencies" $TotalUniqueJobTitles $(DisplayFrequencies "JobTitle" $JobTitles) | Out-File $JobTitles_HTML.Path -Append -Encoding utf8
    #		- #JobTitlesWithoutGroupAssignments
    ExportTo-HTML_NewTable-Div "JobTitlesWithoutGroupAssignments" $TotalJobTitlesWithoutGroupAssignments $(DisplayFrequencies "JobTitle" $JobTitlesWithoutGroupAssignments) | Out-File $JobTitles_HTML.Path -Append -Encoding utf8
    End-InfoDiv              $JobTitles_HTML.Path
    # **Info-Div_2**
    #	- SubTitle: #GroupAssignmentFrequenciesInEachJobTitle
    #	- SubInfo: JobTitles_SubInfo_Div2
    New-InfoDiv              $JobTitles_HTML.Path "GroupAssignmentFrequenciesInEachJobTitle" $JobTitles_SubInfo_Div2
    #	- *Content:*
    #		- #GroupAssignmentFrequenciesInEachJobTitle**
    Get-GroupsInEachPropertyFrequencyObject $JobTitles "JobTitle" $JobTitles_HTML.Path
    End-InfoDiv              $JobTitles_HTML.Path

    Append-HtmlReport_Footer $JobTitles_HTML.Path
    ################################################## #Departments - Info-Div_Count: 2
    # Title: Departments
    # MainHeaderString: $FolderName
    # SubTitle: Departments
    # SubInfo: Departments_SubInfo
    # **Info-Div_1**
    #	- SubTitle: Departments Overview
    #	- SubInfo: Departments_SubInfo_Div1
    New-InfoDiv              $Departments_HTML.Path "DepartmentsOverview" $Departments_SubInfo_Div1
    #	- *Content:*
    #		- #DepartmentFrequencies
    ExportTo-HTML_NewTable-Div "DepartmentFrequencies" $TotalUniqueDepartments $(DisplayFrequencies "Department" $Departments) | Out-File $Departments_HTML.Path -Append -Encoding utf8
    #		- #DepartmentsWithoutGroupAssignments
    ExportTo-HTML_NewTable-Div "DepartmentsWithoutGroupAssignments" $TotalDepartmentsWithoutGroupAssignments $(DisplayFrequencies "Department" $DepartmentsWithoutGroupAssignments) | Out-File $Departments_HTML.Path -Append -Encoding utf8
    End-InfoDiv              $Departments_HTML.Path 
    # **Info-Div_2**
    #	- SubTitle: #GroupAssignmentFrequenciesInEachDepartment
    #	- SubInfo: Departments_SubInfo_Div2
    New-InfoDiv              $Departments_HTML.Path "GroupAssignmentFrequenciesInEachDepartment" $Departments_SubInfo_Div2
    #	- *Content:*
    #		- #GroupAssignmentFrequenciesInEachDepartment**
    Get-GroupsInEachPropertyFrequencyObject $Departments "Department" $Departments_HTML.Path
    End-InfoDiv              $Departments_HTML.Path 

    Append-HtmlReport_Footer $Departments_HTML.Path

    Write-Output "Your Report is Ready and can be found @: "
    $ReportFolder.Path

    Write-Output "`nProgram completed @ $(date)"
    Write-Host "Program completed @ $(date)"
}

main

##########################################################################################################################
# $TotalGroupsWithUsers_ButWithoutMembers = ($Groups | where{($_.Departments -ne "NULL") -and ($_.GroupMembers -eq $null)}).count    #0   Expected to be 0 - Cannot have users in group without members

# $TotalGroupsWithMembers    = ($Groups | where{$_.GroupMembers -ne $null}).count   
# $TotalGroupsWithoutMembers = ($Groups | where{$_.GroupMembers -eq $null}).count   
# $TotalGroupsWithoutUsers   = ($Groups | where{$_.Departments -eq "NULL"}).count   
# $TotalGroupsWithUsers      = ($Groups | where{$_.Departments -ne "NULL"}).count
# $TotalGroupsWithoutUsers_ButWithMembers = ($Groups | where{($_.Departments -eq "NULL") -and ($_.GroupMembers -ne $null)}).count

# TotalGroupsWithoutUsers - TotalGroupsWithoutMembers = TotalGroupsWithoutUsers_ButWithMembers
# TotalGroupsWithMembers  - TotalGroupsWithUsers      = TotalGroupsWithoutUsers_ButWithMembers
# TotalGroupsWithUsers    + TotalGroupsWithoutUsers   = TotalGroups
# (TotalGroupsWithUsers + TotalGroupsWithoutUsers_ButWithMembers = TotalGroupsWithMembers) + TotalGroupsWithoutMembers = TotalGroups
# TotalUniqueGroupNames = (GroupAssignmentFrequencies).count
##########################################################################################################################

#Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "Departments" "Department"
#Get-FrequencyObjectInEachGroup $Groups $GroupAssignmentFrequencies "JobTitles" "JobTitle"
#Get-GroupsInEachPropertyFrequencyObject $Departments "Department"
#Get-GroupsInEachPropertyFrequencyObject $JobTitles "JobTitle" 

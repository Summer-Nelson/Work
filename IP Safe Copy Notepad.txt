#Create New User in Active Directory

<#
Notes: 
    MUST OPEN POWERSHELL AS ADMIN
		 
	NEED TO HAVE EXCHANGE RECIPIENT MANAGEMENT TOOLS INSTALLED:
        https://learn.microsoft.com/en-us/exchange/plan-and-deploy/post-installation-tasks/install-management-tools?view=exchserver-2019
        https://learn.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools

    NEED TO HAVE AZURE AD INSTALLED TO RUN:
	    Install-Module -Name "PnP.PowerShell"
	    Install-Module AzureAD
	    Import-Module AzureAD

	ADDING A BRANCH:
		Add the name of the branch to the global:branchList (AT THE END OF THE ARRAY).
		Create a new department function.
		Add a new elseif in the edit properties function.
		Add a new elseif to userSetup function with with the correct path to the branch.
		Change/enter the correct paths for users and groups in the userList, groupsList, and managerList variables.
		Follow the instructions on how to add the locations to this branch (see below).
		
	ADDING A LOCATION TO A BRANCH:
		Create a new array variable [ $_ = @() ] named after the city it is in.
		Add the list of all departments in that locations array (AT THE END OF THE ARRAY).
		Add the variable to the allDepartments_ array.
		After adding the variable, go to the corresponding if statment (based on branch location) under the editProperties function.
		In the editProperties function, add the office, company, city, street, zip, country, and region information  to their corresponding array variables (AT THE END OF THE ARRAY).

	ADDING A DEPARTMENT:
		Add the name of the department to the end of the deparment list of the corresponding location.
#>

Add-Type -AssemblyName System.Windows.Forms

$global:cancelFlag = $null

$global:branchList = @(‘Branch1’,‘Branch2’,‘Branch3’) 
$global:userList = $null
$global:deptList = $null
$global:managerList = $null
$global:officeList = $null 
$global:companyList = $null 
$global:cityList = $null
$global:streetList = $null 
$global:zipList = $null
$global:regionList = $null
$global:countryList = $null
$global:groupsList = $null

$global:selectedBranch = $null
$global:selectedUser = $null
$global:selectedManager = $null
$global:selectedOffice = $null
$global:selectedDept = $null
$global:selectedTitle = $null
$global:selectedCompany = $null
$global:selectedStreet = $null
$global:selectedCity = $null
$global:selectedZip = $null
$global:selectedRegion = $null
$global:selectedCountry = $null
$global:selectedGroups = $null

$global:userFirstName = $null
$global:userLastName = $null #Surname
$global:fullName = $null #CN
$global:displayName = $null #DisplayName
$global:mailOption = $null #SamAccountName
$global:globalMail = $null
$global:primaryMail = $null #UserPrincipalName



#Department Information
<#
Function: departmentsBranch1
Purpose: Create and manage the list of Departments for each location of the Branch1 branches.
Parameters: None
Variables: Location1- Array of all the departments at the Location1 location
           Location2 - Array of all the departments at the Location2 location
           allDepartmentsBranch1 - An array that contains arrays of each depatment at each location. Stored and navigated by index.
#>
function departmentsBranch1(){
    $Location1 = @("L1 D1 Branch1", "L1 D2 Branch1")  #add real list  
    $Location2 = @("L2 D1 Branch1", "L2 D2 Branch1")  #add real list

    $allDepartmentsBranch1 = @($Location1, $Location2) 
        
    return $allDepartmentsBranch1
}

<#
Function: departmentsBranch2
Purpose: Create and manage the list of Departments for each location of the Branch1 branches.
Parameters: None
Variables: Location1- Array of all the departments at the Location1 location
           Location2 - Array of all the departments at the Location2 location
		   Location3 - Array of all the departments at the Location3 location
           allDepartmentsBranch2 - An array that contains arrays of each depatment at each location. Stored and navigated by index.
#>
function departmentsBranch2(){
    $Location1 = @("L1 D1 Branch2", "L1 D2 Branch2")  #add real list  
    $Location2 = @("L2 D1 Branch2", "L2 D2 Branch2")  #add real list
	$Location3 = @("L3 D1 Branch2", "L3 D2 Branch2")  #add real list
	
    $allDepartmentsBranch2 = @($Location1, $Location2,$Location3) 
        
    return $allDepartmentsBranch2
}

<#
Function: departmentsBranch3
Purpose: Create and manage the list of Departments for each location of the Branch1 branches.
Parameters: None
Variables: Location1- Array of all the departments at the Location1 location
           Location2 - Array of all the departments at the Location2 location
		   Location3 - Array of all the departments at the Location3 location
		   Location4 - Array of all the departments at the Location4 location
           allDepartmentsBranch3 - An array that contains arrays of each depatment at each location. Stored and navigated by index.
#>
function departmentsBranch3(){
    $Location1 = @("L1 D1 Branch3", "L1 D2 Branch3")  #add real list  
    $Location2 = @("L2 D1 Branch3", "L2 D2 Branch3")  #add real list
	$Location3 = @("L3 D1 Branch3", "L3 D2 Branch3")  #add real list
	$Location4 = @("L4 D1 Branch3", "L4 D2 Branch3")  #add real list

    $allDepartmentsBranch3 = @($Location1, $Location2,$Location3,$Location4) 
        
    return $allDepartmentsBranch3
}



#Setup Forms
<#
Name: branchForm
Purpose: Create form for the selection of the users branch
Parameters: None
Variables: formBranch - Creates the form
           ListBranch - creats the branch dropdown box
           DescriptionBranch - Creates the "Branch" label
           global:selectedBranch - Sets the branch that is selected by the user 
           okButton - Makes an OK button
           cancelButton - Makes a Cancel button
           result - Sets what is currenly selected in the dropdown box
           global:cancelFlag - Boolean yes or no for weter or not to end the program.
#>
function branchForm(){
    #Blank form
    $formBranch = New-Object Windows.Forms.Form
    $formBranch.Size = New-Object Drawing.Size @(800,600)
    $formBranch.Text = "Create New User"

    #Make dropdown list
    $ListBranch = New-Object system.Windows.Forms.ComboBox
    $ListBranch.text = “”
    $ListBranch.width = 200
    $ListBranch.autosize = $true

    # Add the items in the dropdown list
    $global:branchList | ForEach-Object {[void] $ListBranch.Items.Add($_)}
    # Select the default value
    #$List.SelectedIndex = 0
    $ListBranch.Text = "-- Select Branch --"
    $ListBranch.location = New-Object System.Drawing.Point(70,100)
    $ListBranch.Font = ‘Microsoft Sans Serif,10’
    $formBranch.Controls.Add($ListBranch)

    #Add a label
    $DescriptionBranch = New-Object system.Windows.Forms.Label
    $DescriptionBranch.Text = “Select Branch:”
    $DescriptionBranch.AutoSize = $false
    $DescriptionBranch.width = 450
    $DescriptionBranch.height = 50
    $DescriptionBranch.location = New-Object System.Drawing.Point(20,50)
    $DescriptionBranch.Font = ‘Microsoft Sans Serif,10’
    $formBranch.Controls.Add($DescriptionBranch)

    #Catch changes to the list
    $ListBranch.add_SelectedIndexChanged({
        $global:selectedBranch = $global:branchList[$ListBranch.SelectedIndex]
    })
    
    #Ok Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,150)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formBranch.AcceptButton = $okButton
    $formBranch.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(175,150)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formBranch.CancelButton = $cancelButton
    $formBranch.Controls.Add($cancelButton)

    $result = $formBranch.ShowDialog()
    if($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $global:cancelFlag = 1
    }

    #Dispose unused variables
    $formBranch.Dispose()
}

<#
Name: propertiesForm
Purpose: Create form for the selection of the users properties
Parameters: None
Variables: formProperties - Creates the properties form.
           
           DescriptionCopy - Creates the "User to Copy" label.
           DescriptionTitle - Creates the "Title" label.
           DescriptionDept - Creates the "Department" label.
           DescriptionManager - Creates the "Manager" label.
           DescriptionOffice - Creates the "Office" label.
           DescriptionCompany - Creates the "Company" label.
           DescriptionCity - Creates the "City" label.
           DescriptionStreet - Creates the "Street" label.
           DescriptionZip - Creates the "Zip" label.
           DescriptionCountry - Creates the "Country" label.
           DescriptionRegion - Creates the "Region" label.
           
           ListCopy - Create the drop down box that contains the list of available users.
           ListOffice - Create the drop down box that contains the list of offices.
           ListDept - Create the drop down box that contains the list of departments.
           ListManager - Create the drop down box that contains the list of managers.
           
           TextCompany - Creates a label that displays the corresponding company of the selected office.
           TextStreet - Creates a label that displays the corresponding street address of the selected office.
           TextCity - Creates a label that displays the corresponding city of the selected office.
           TextZip - Creates a label that displays the corresponding zip of the selected office.
           TextRegion - Creates a label that displays the corresponding region of the selected office.
           TextCountry - Creates a label that displays the corresponding country of the selected office.
           
           global:userList - A string array that contains the list of all users under the selected branch.
           global:officeList - A string array that contains the list of all the offices under the selected branch.
           global:managerList - A string array that contains the list of all users whos titles include "Manager" under the selected branch.
           global:deptList - A string array that contains the list of all users under the selected office.
           
           global:selectedUser - Stores a user object that is the same as the selected user to be copied.
           global:selectedManager - Stores a user object that is the same as the selected manager.
           global:selectedOffice - Stores a string that is the same as the office of the selected user or that was chosen.
           global:selectedDept - Stores a string that is the same as the department of the selected user or that was chosen.
           global:selectedCompany - Stores a string that is the same as the company of the selected user or that was chosen.
           global:selectedStreet - Stores a string that is the same as the corresponding street of the selected office.
           global:selectedCity - Stores a string that is the same as the corresponding city of the selected office.
           global:selectedZip - Stores a string that is the same as the corresponding zip code of the selected office.
           global:selectedRegion - Stores a string that is the same as the corresponding region of the selected office.
           global:selectedCountry -  Stores a string that is the same as the corresponding country of the selected office.
           global:selectedTitle - Stores a string that is the same as the corresponding title of the selected office.

           details - Stores all the properties of the selected user.
           flag - Checks if the properties of the selected user is the same as a property in the corresponding __List. 1 is yes, 0 is no.
           path - Stores the path of the selected manager.
           Manager - Gets the user object at the specified path.
           titleTextBox - Creates a text box that the tile will be typed into.
           result - Checks if the reult of the form was okay or cancel.

           okButton - Makes an OK button.
           cancelButton - Makes a Cancel button.
           global:cancelFlag - Boolean yes or no for weter or not to end the program.
#>
function propertiesForm(){
    #Blank form
    $formProperties = New-Object Windows.Forms.Form
    $formProperties.Size = New-Object Drawing.Size @(800,800)
    $formProperties.Text = "Create New User"
     
    #Add user label
    $DescriptionCopy = New-Object system.Windows.Forms.Label
    $DescriptionCopy.text = “User to Copy: ”
    $DescriptionCopy.AutoSize = $false
    $DescriptionCopy.width = 150
    $DescriptionCopy.height = 50
    $DescriptionCopy.location = New-Object System.Drawing.Point(20,50)
    $DescriptionCopy.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionCopy)
    
    #Make dropdown list
    $ListCopy = New-Object system.Windows.Forms.ComboBox
    $ListCopy.text = “-- Pick user to copy --”
    $ListCopy.width = 250
    $ListCopy.AutoSize = $true
    $ListCopy.location = New-Object System.Drawing.Point(180,50)
    $ListCopy.Font = ‘Microsoft Sans Serif,12’
    $formProperties.Controls.Add($ListCopy)

    #Get user properties
    editProperties($global:selectedBranch)
    
    #Fill dropdown list
    $global:userList | forEach-Object {$ListCopy.Items.Add($_.Name)}

    #Properties lables
    $DescriptionTitle = New-Object system.Windows.Forms.Label
    $DescriptionTitle.text = “Title: ”
    $DescriptionTitle.AutoSize = $false
    $DescriptionTitle.width = 150
    $DescriptionTitle.height = 50
    $DescriptionTitle.location = New-Object System.Drawing.Point(50,300)
    $DescriptionTitle.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionTitle)

    $DescriptionDept = New-Object system.Windows.Forms.Label
    $DescriptionDept.text = “Department:  ”
    $DescriptionDept.AutoSize = $false
    $DescriptionDept.width = 150
    $DescriptionDept.height = 50
    $DescriptionDept.location = New-Object System.Drawing.Point(50,200)
    $DescriptionDept.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionDept)

    $DescriptionManager = New-Object system.Windows.Forms.Label
    $DescriptionManager.text = “Manager: ”
    $DescriptionManager.AutoSize = $false
    $DescriptionManager.width = 150
    $DescriptionManager.height = 50
    $DescriptionManager.location = New-Object System.Drawing.Point(50,250)
    $DescriptionManager.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionManager)

    $DescriptionOffice = New-Object system.Windows.Forms.Label
    $DescriptionOffice.text = “Office: ”
    $DescriptionOffice.AutoSize = $false
    $DescriptionOffice.width = 100
    $DescriptionOffice.height = 50
    $DescriptionOffice.location = New-Object System.Drawing.Point(50,150)
    $DescriptionOffice.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionOffice)

    $DescriptionCompany = New-Object system.Windows.Forms.Label
    $DescriptionCompany.text = “Company: ”
    $DescriptionCompany.AutoSize = $false
    $DescriptionCompany.width = 150
    $DescriptionCompany.height = 50
    $DescriptionCompany.location = New-Object System.Drawing.Point(50,350)
    $DescriptionCompany.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionCompany)

    $DescriptionCity = New-Object system.Windows.Forms.Label
    $DescriptionCity.text = “City: ”
    $DescriptionCity.AutoSize = $false
    $DescriptionCity.width = 150
    $DescriptionCity.height = 50
    $DescriptionCity.location = New-Object System.Drawing.Point(50,450)
    $DescriptionCity.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionCity)

    $DescriptionStreet = New-Object system.Windows.Forms.Label
    $DescriptionStreet.text = “Street: ”
    $DescriptionStreet.AutoSize = $false
    $DescriptionStreet.width = 150
    $DescriptionStreet.height = 50
    $DescriptionStreet.location = New-Object System.Drawing.Point(50,400)
    $DescriptionStreet.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionStreet)

    $DescriptionZip = New-Object system.Windows.Forms.Label
    $DescriptionZip.text = “Zip: ”
    $DescriptionZip.AutoSize = $false
    $DescriptionZip.width = 150
    $DescriptionZip.height = 50
    $DescriptionZip.location = New-Object System.Drawing.Point(50,500)
    $DescriptionZip.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionZip)

    $DescriptionCountry = New-Object system.Windows.Forms.Label
    $DescriptionCountry.text = “Region: ”
    $DescriptionCountry.AutoSize = $false
    $DescriptionCountry.width = 160
    $DescriptionCountry.height = 50
    $DescriptionCountry.location = New-Object System.Drawing.Point(50,550)
    $DescriptionCountry.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionCountry)

    $DescriptionRegion = New-Object system.Windows.Forms.Label
    $DescriptionRegion.text = “Country: ”
    $DescriptionRegion.AutoSize = $false
    $DescriptionRegion.width = 170
    $DescriptionRegion.height = 50
    $DescriptionRegion.location = New-Object System.Drawing.Point(50,600)
    $DescriptionRegion.Font = ‘Microsoft Sans Serif,10’
    $formProperties.Controls.Add($DescriptionRegion)

    #Office dropdown
    $ListOffice = New-Object system.Windows.Forms.ComboBox
    $ListOffice.width = 250
    $ListOffice.autosize = $true
    $ListOffice.Text = "--Select Office--"

    $ListDept = New-Object system.Windows.Forms.ComboBox
    $ListDept.width = 300
    $ListDept.autosize = $true
    $ListDept.Text = "--Select Department--"

    $ListManager = New-Object system.Windows.Forms.ComboBox
    $ListManager.width = 300
    $ListManager.autosize = $true
    $ListManager.Text = "--Select Manager--"

    $TextCompany = New-Object system.Windows.Forms.Label
    $TextCompany.width = 200
    $TextCompany.autosize = $true

    $TextStreet = New-Object system.Windows.Forms.Label
    $TextStreet.width = 200
    $TextStreet.autosize = $true

    $TextCity = New-Object system.Windows.Forms.Label
    $TextCity.width = 200
    $TextCity.autosize = $true

    $TextZip = New-Object system.Windows.Forms.Label
    $TextZip.width = 200
    $TextZip.autosize = $true

    $TextRegion = New-Object system.Windows.Forms.Label
    $TextRegion.width = 200
    $TextRegion.autosize = $true

    $TextCountry = New-Object system.Windows.Forms.Label
    $TextCountry.width = 200
    $TextCountry.autosize = $true

    #Add the items in the dropdown list
    $global:officeList | ForEach-Object {[void] $ListOffice.Items.Add($_)}

    #Select the default value
    $ListOffice.location = New-Object System.Drawing.Point(180,150)
    $ListOffice.Font = ‘Microsoft Sans Serif,12’
    $formProperties.Controls.Add($ListOffice)

    $ListManager.location = New-Object System.Drawing.Point(250,250)
    $ListManager.Font = ‘Microsoft Sans Serif,12’
    $formProperties.Controls.Add($ListManager)

    $global:managerList | ForEach-Object {[void] $ListManager.Items.Add($_.Name)}

    

    #COPY USER INFO
    $ListCopy.add_SelectedIndexChanged({
        $global:selectedUser = $ListCopy.SelectedItem

        #GetUserDetails
        $details = Get-ADUser -Filter "Name -eq '$global:selectedUser'" -Properties *
        
        #Match Manager
        $ListManager.SelectedIndex = 0
        $flag = 0

        $path = $details.manager
        $Manager = Get-ADUser -Identity "$path" -Properties *
        
        $global:managerList | ForEach-Object{

            if($Manager.Name -eq $_){
                $global:selectedManager = $Manager
                $flag = 1
            }
            elseif(($ListManager.SelectedIndex -eq ($global:managerList.Length-1)) -and ($flag -eq 0)){
                $global:managerList += $Manager.Name
                        
                $ListManager.Items.Add($Manager.Name)
                $ListManager.SelectedIndex ++
                
                #return CN not name       
                $global:selectedManager = $Manager
                $flag = 1
            }
            elseif($flag -eq 0){
                $ListManager.SelectedIndex ++
            }
        }

        #Match Office
        $ListOffice.SelectedIndex = 0
        $flag = 0

        $global:officeList | ForEach-Object{
            if($details.office -eq $_){
                $global:selectedOffice = $officeList[$ListOffice.SelectedIndex]
                $flag = 1
            }
            elseif(($ListOffice.SelectedIndex -eq ($global:officeList.Length-1)) -and ($flag -eq 0)){
                $global:officeList += $details.office
                        
                $ListOffice.Items.Add($details.office)
                $ListOffice.SelectedIndex ++

                #Copy Address Information

                $global:companyList += $details.Company
                $global:streetList += $details.StreetAddress
                $global:cityList  += $details.City
                $global:zipList += $details.PostalCode
                $global:regionList += $details.State
                $global:countryList += $details.Country

                 # Add the Company
                $TextCompany.text = $companyList[$ListOffice.SelectedIndex]       
                $TextCompany.location = New-Object System.Drawing.Point(250,350)
                $TextCompany.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextCompany)
                $global:selectedCompany = $global:companyList[$ListOffice.SelectedIndex]  

                # Add the Street
                $TextStreet.text = $streetList[$ListOffice.SelectedIndex]       
                $TextStreet.location = New-Object System.Drawing.Point(250,400)
                $TextStreet.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextStreet)
                $global:selectedStreet = $global:streetList[$ListOffice.SelectedIndex]

                # Add the City
                $TextCity.text = $cityList[$ListOffice.SelectedIndex]       
                $TextCity.location = New-Object System.Drawing.Point(250,450)
                $TextCity.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextCity)
                $global:selectedCity = $global:cityList[$ListOffice.SelectedIndex]

                # Add the Zip
                $TextZip.text = $zipList[$ListOffice.SelectedIndex]       
                $TextZip.location = New-Object System.Drawing.Point(250,500)
                $TextZip.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextZip)
                $global:selectedZip = $global:zipList[$ListOffice.SelectedIndex]

                # Add the Region
                $TextRegion.text = $regionList[$ListOffice.SelectedIndex]       
                $TextRegion.location = New-Object System.Drawing.Point(250,550)
                $TextRegion.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextRegion)
                $global:selectedRegion = $global:regionList[$ListOffice.SelectedIndex]

                # Add the Country
                $TextCountry.text = $countryList[$ListOffice.SelectedIndex]       
                $TextCountry.location = New-Object System.Drawing.Point(250,600)
                $TextCountry.Font = ‘Microsoft Sans Serif,10’
                $formProperties.Controls.Add($TextCountry)
                $global:selectedCountry = $global:countryList[$ListOffice.SelectedIndex]

                        
                        
                $ListDept.Items.Add($details.department)
                $ListDept.SelectedIndex ++

                $global:deptList += $ListDept.Text

                $global:selectedDept = $ListDept.Text
       
                $global:selectedOffice = $officeList[$ListOffice.SelectedIndex]
                $flag = 1
            }
            elseif($flag -eq 0){
                $ListOffice.SelectedIndex ++
            }
        }

        #Match Department
        $ListDept.SelectedIndex = 0
        $flag = 0

        $global:deptList[$ListOffice.SelectedIndex] | ForEach-Object{
            if($details.department -eq $_){
                $global:selectedDept = $ListDept.Text
                $flag = 1
            }
            elseif(($ListDept.SelectedIndex -eq ($global:deptList[$ListOffice.SelectedIndex].Length-1)) -and ($flag -eq 0)){
                #$global:deptList += $ListDept.Text
                        
                $ListDept.Items.Add($details.department)
                $ListDept.SelectedIndex ++

                $global:deptList += $ListDept.Text
                #$global:deptList.Items.Add($details.department)         

                $global:selectedDept = $ListDept.Text
                $flag = 1
            }
            elseif($flag -eq 0){
                $ListDept.SelectedIndex ++
            }
        }
               
    })


    #Catch changes to manager list
    $ListManager.add_SelectedIndexChanged({
        $global:selectedManager = $managerList[$ListManager.SelectedIndex]
    })

    #Catch changes to office list
    $ListOffice.add_SelectedIndexChanged({
        $global:selectedOffice = $ListOffice.SelectedIndex

        # Add the items in the dropdown list
        $ListDept.Items.Clear()
        $global:deptList[$selectedOffice] | ForEach-Object {[void] $ListDept.Items.Add($_)}
         
        $ListDept.location = New-Object System.Drawing.Point(250,200)
        $ListDept.Font = ‘Microsoft Sans Serif,12’
        $formProperties.Controls.Add($ListDept)
        
        #Catch changes to the Dept list
        $ListDept.add_SelectedIndexChanged({
            $global:selectedDept = $ListDept.Text
         })

        # Add the Company
        $TextCompany.text = $companyList[$selectedOffice]       
        $TextCompany.location = New-Object System.Drawing.Point(250,350)
        $TextCompany.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextCompany)
        $global:selectedCompany = $companyList[$selectedOffice]  

        # Add the Street
        $TextStreet.text = $streetList[$selectedOffice]       
        $TextStreet.location = New-Object System.Drawing.Point(250,400)
        $TextStreet.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextStreet)
        $global:selectedStreet = $streetList[$selectedOffice]

        # Add the City
        $TextCity.text = $cityList[$selectedOffice]       
        $TextCity.location = New-Object System.Drawing.Point(250,450)
        $TextCity.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextCity)
        $global:selectedCity = $cityList[$selectedOffice]

        # Add the Zip
        $TextZip.text = $zipList[$selectedOffice]       
        $TextZip.location = New-Object System.Drawing.Point(250,500)
        $TextZip.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextZip)
        $global:selectedZip = $zipList[$selectedOffice]

        # Add the Region
        $TextRegion.text = $regionList[$selectedOffice]       
        $TextRegion.location = New-Object System.Drawing.Point(250,550)
        $TextRegion.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextRegion)
        $global:selectedRegion = $regionList[$selectedOffice]

        # Add the Country
        $TextCountry.text = $countryList[$selectedOffice]       
        $TextCountry.location = New-Object System.Drawing.Point(250,600)
        $TextCountry.Font = ‘Microsoft Sans Serif,10’
        $formProperties.Controls.Add($TextCountry)
        $global:selectedCountry = $countryList[$selectedOffice]

        $global:selectedOffice = $officeList[$ListOffice.SelectedIndex]
    })

    #Ok Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250,650)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formProperties.AcceptButton = $okButton
    $formProperties.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350,650)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formProperties.CancelButton = $cancelButton
    $formProperties.Controls.Add($cancelButton)
 
    #Add title
    $titleTextBox = New-Object System.Windows.Forms.TextBox
    $titleTextBox.location = New-Object System.Drawing.Point(250,300)
    $titleTextBox.Width = 300
    $titleTextBox.Font = ‘Microsoft Sans Serif,15’
    $formProperties.Controls.Add($titleTextBox)

    if($global:selectedUser -ne $null){
        $titleTextBox.Text = $details.Title
    }

    $formProperties.Add_Shown({$titleTextBox.Select()})
    $result = $formProperties.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        $global:selectedTitle = $titleTextBox.Text
    }  

    if($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $global:cancelFlag = 1
    }

    $formProperties.Dispose()
}

<#
Name: nameForm
Purpose: Create form for the selection of the users name
Parameters: None
Variables: formName - Creates the form
           okButton - Makes an OK button
           cancelButton - Makes a Cancel button
           DescriptionFirst - Creates the "First Name" label
           DescriptionLast - Creates the "Lanst Name" label
           fistNameTextBox - Makes the text box where you can enter the users first name
           lastNameTextBox - Makes the text box where you can enter the users last name
           result - Sets what is currently typed in the textbox
           global:userFirstName - String variable that sets the users first name.
           global:userLastName - String variable that sets the users last name.
#>
function nameForm(){
    #Blank form
    $formName = New-Object Windows.Forms.Form
    $formName.Size = New-Object Drawing.Size @(800,600)
    $formName.Text = "Create New User"

    #Ok Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250,450)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formName.AcceptButton = $okButton
    $formName.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350,450)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formName.CancelButton = $cancelButton
    $formName.Controls.Add($cancelButton)

    #Add a label
    $DescriptionFirst = New-Object system.Windows.Forms.Label
    $DescriptionFirst.text = “First Name:”
    $DescriptionFirst.AutoSize = $false
    $DescriptionFirst.width = 150
    $DescriptionFirst.height = 50
    $DescriptionFirst.location = New-Object System.Drawing.Point(20,100)
    $DescriptionFirst.Font = ‘Microsoft Sans Serif,10’
    $formName.Controls.Add($DescriptionFirst)

    #Add a label
    $DescriptionLast = New-Object system.Windows.Forms.Label
    $DescriptionLast.text = “Last Name:”
    $DescriptionLast.AutoSize = $false
    $DescriptionLast.width = 150
    $DescriptionLast.height = 50
    $DescriptionLast.location = New-Object System.Drawing.Point(20,175)
    $DescriptionLast.Font = ‘Microsoft Sans Serif,10’
    $formName.Controls.Add($DescriptionLast)

    #Add First Name
    $fistNameTextBox = New-Object System.Windows.Forms.TextBox
    $fistNameTextBox.location = New-Object System.Drawing.Point(190,100)
    $fistNameTextBox.Width = 200
    $fistNameTextBox.Font = ‘Microsoft Sans Serif,15’
    $formName.Controls.Add($fistNameTextBox)

    $formName.Add_Shown({$fistNameTextBox.Select()})

    #Add Last Name
    $lastNameTextBox = New-Object System.Windows.Forms.TextBox
    $lastNameTextBox.location = New-Object System.Drawing.Point(190,175)
    $lastNameTextBox.Width = 200
    $lastNameTextBox.Font = ‘Microsoft Sans Serif,15’
    $formName.Controls.Add($lastNameTextBox)

    $formName.Add_Shown({$lastNameTextBox.Select()})

    #Output and return
    $result = $formName.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        $global:userFirstName = $fistNameTextBox.Text
        $global:userLastName = $lastNameTextBox.Text
    }
    if($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $global:cancelFlag = 1
    }

    $formName.Dispose()
}

<#
Name: groupsForm
Purpose: Creates the form to choose the users groups
Paramaters: None
Variables: formGroups - Creates the form
           DescriptionGroups - Creates the "Groups" label
           listBox - Creates the list box that will hold the groups.
           checkboxes - Keeps track of what is checked/selected and not
           okButton - Makes an OK button
           cancelButton - Makes a Cancel button
           details - Holds the properties of the selected/copied user
           copiedGroups - String array that contail a list of all the groups the copied user was a part of.
           index - Keeps track of the index of the copied groups array ($copiedGroups)
           result - Sets if a responce is OK or Cancel
           global:selectedGroups - String array variable that sets the users first name.
           global:cancelFlag - Boolean yes or no for weter or not to end the program.
#>
function groupsForm(){
    #Blank form
    $formGroups = New-Object Windows.Forms.Form
    $formGroups.Size = New-Object Drawing.Size @(800,600)
    $formGroups.Text = "Create New User"

    #Add a label
    $DescriptionGroups = New-Object system.Windows.Forms.Label
    $DescriptionGroups.text = “Groups:”
    $DescriptionGroups.AutoSize = $false
    $DescriptionGroups.width = 120
    $DescriptionGroups.height = 40
    $DescriptionGroups.location = New-Object System.Drawing.Point(20,30)
    $DescriptionGroups.Font = ‘Microsoft Sans Serif,10’
    $formGroups.Controls.Add($DescriptionGroups)

    #Add listbox
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(30,75)
    $listBox.Height = 400
    $listBox.Width = 700
    $listBox.Font = ‘Microsoft Sans Serif,12’
    $listBox.SelectionMode = 'MultiSimple'
    $checkboxes = foreach($item in $global:groupsList){
        [void] $listBox.Items.Add($item.Name)
    }

    $formGroups.Controls.Add($listBox)

    #Ok Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250,500)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formGroups.AcceptButton = $okButton  
    $formGroups.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350,500)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formGroups.CancelButton = $cancelButton
    $formGroups.Controls.Add($cancelButton)

    #Copy Groups of Selected User
    if($global:selectedUser -ne $null){
        $details = Get-ADUser -Filter "Name -eq '$global:selectedUser'" -Properties *

        $copiedGroups = @(Get-ADPrincipalGroupMembership -Identity $details | Select name)
        
        $copiedGroups | ForEach-Object{ 
            $index = $global:groupsList.Name.IndexOf($_.Name)
            $listBox.SetSelected($index, $True)
        }
    }
    
    #Selecting Groups
    $result = $formGroups.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK){
        $global:selectedGroups = $listBox.SelectedItems 
    }
    if($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $global:cancelFlag = 1
    }
}

<#
Name: emailForm
Purpose: Creates the form to manually set the email.
Paramaters: None
Variables: formEmail - Creates the form
           okButton - Makes an OK button.
           cancelButton - Makes a Cancel button.
           DescriptionError - Creates an error label/message for the emailForm.
           DescriptionFirstEmail - Creates the "Please enter new - email username" label.
           DescriptionLastEmail - Creates the "Confirm new email - username" label.
           DescriptionEmail1 - Creates the “@email.com” label.
           DescriptionEmail2 - Creates the “@email.com” label.
           fistEmailTextBox - Creates the textbox for the first email.
           lastEmailTextBox - Creates the textbox for the confirm email.
           result - Sets the responce/dialoge of the emailForm.
           global:cancelFlag - Boolean yes or no for weter or not to end the program.
#>
function emailForm(){
    #Blank form
    $formEmail = New-Object Windows.Forms.Form
    $formEmail.Size = New-Object Drawing.Size @(800,600)
    $formEmail.Text = "Create New User"

    #Ok Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(250,450)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formEmail.AcceptButton = $okButton
    $formEmail.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350,450)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formEmail.CancelButton = $cancelButton
    $formEmail.Controls.Add($cancelButton)

    #Add a label
    $DescriptionError = New-Object system.Windows.Forms.Label
    $DescriptionError.text = “Error, this email already exists or the entered emails did not match”
    $DescriptionError.AutoSize = $false
    $DescriptionError.width = 650
    $DescriptionError.height = 50
    $DescriptionError.location = New-Object System.Drawing.Point(100,30)
    $DescriptionError.Font = ‘Microsoft Sans Serif,10’
    $DescriptionError.ForeColor = "Red"
    $formEmail.Controls.Add($DescriptionError)

    #Add a label
    $DescriptionFirstEmail = New-Object system.Windows.Forms.Label
    $DescriptionFirstEmail.text = “Please enter new - email username”
    $DescriptionFirstEmail.AutoSize = $false
    $DescriptionFirstEmail.width = 200
    $DescriptionFirstEmail.height = 50
    $DescriptionFirstEmail.location = New-Object System.Drawing.Point(20,100)
    $DescriptionFirstEmail.Font = ‘Microsoft Sans Serif,10’
    $formEmail.Controls.Add($DescriptionFirstEmail)

    #Add a label
    $DescriptionLastEmail = New-Object system.Windows.Forms.Label
    $DescriptionLastEmail.text = “Confirm new email - username”
    $DescriptionLastEmail.AutoSize = $false
    $DescriptionLastEmail.width = 200
    $DescriptionLastEmail.height = 50
    $DescriptionLastEmail.location = New-Object System.Drawing.Point(20,175)
    $DescriptionLastEmail.Font = ‘Microsoft Sans Serif,10’
    $formEmail.Controls.Add($DescriptionLastEmail)

    #Add a label
    $DescriptionEmail1 = New-Object system.Windows.Forms.Label
    $DescriptionEmail1.text = “@email.com”
    $DescriptionEmail1.AutoSize = $false
    $DescriptionEmail1.width = 250
    $DescriptionEmail1.height = 50
    $DescriptionEmail1.location = New-Object System.Drawing.Point(450,100)
    $DescriptionEmail1.Font = ‘Microsoft Sans Serif,10’
    $formEmail.Controls.Add($DescriptionEmail1)

    #Add a label
    $DescriptionEmail2 = New-Object system.Windows.Forms.Label
    $DescriptionEmail2.text = “@email.com”
    $DescriptionEmail2.AutoSize = $false
    $DescriptionEmail2.width = 250
    $DescriptionEmail2.height = 50
    $DescriptionEmail2.location = New-Object System.Drawing.Point(450,175)
    $DescriptionEmail2.Font = ‘Microsoft Sans Serif,10’
    $formEmail.Controls.Add($DescriptionEmail2)

    #Add Email
    $fistEmailTextBox = New-Object System.Windows.Forms.TextBox
    $fistEmailTextBox.location = New-Object System.Drawing.Point(250,100)
    $fistEmailTextBox.Width = 200
    $fistEmailTextBox.Font = ‘Microsoft Sans Serif,15’
    $formEmail.Controls.Add($fistEmailTextBox)

    $formEmail.Add_Shown({$fistEmailTextBox.Select()})

    #Check Email
    $lastEmailTextBox = New-Object System.Windows.Forms.TextBox
    $lastEmailTextBox.location = New-Object System.Drawing.Point(250,175)
    $lastEmailTextBox.Width = 200
    $lastEmailTextBox.Font = ‘Microsoft Sans Serif,15’
    $formEmail.Controls.Add($lastEmailTextBox)

    $formEmail.Add_Shown({$lastEmailTextBox.Select()})

    #Output and return
    $result = $formEmail.ShowDialog()
    if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and ($fistEmailTextBox.Text -eq $lastEmailTextBox.Text)){
        $global:mailOption = $fistEmailTextBox.Text
    }
    if(($result -eq [System.Windows.Forms.DialogResult]::OK)-and (($fistEmailTextBox.Text -eq "") -or ($lastEmailTextBox.Text -eq "") -or ($fistEmailTextBox.Text -ne $lastEmailTextBox.Text))){
        emailForm
    }
    if($result -eq [System.Windows.Forms.DialogResult]::Cancel){
        $global:cancelFlag = 1
    }

    $formEmail.Dispose()
}



#Background Work
<#
Name: editProperties
Purpose:Set up the elements in the gui dropdown boxes based on branch, user, and office.
Parameters: selectedBranch. What this does is decide what branch the list of users, managers, groups, and offices will be dispalayed
Variables: global:userList - List of users that are a part of the selected branch.
           global:groupsList - List of groups that are a part of the selected branch.
           global:managerList - List of manager that are a part of the selected branch.
           global:deptList - List of departments that are a part of the selected branch.
           global:officeList - List of users that are a part of the selected branch.
           global:companyList - Company name assosiated with selected office.
           global:streetList - Street name assosiated with selected office.
           global:zipList - Zip/area code bof assosiated with selected office.
           global:countryList - Country name assosiated with selected office.
           global:regionList - Region name assosiated with selected office.
#>
function editProperties(){
    if($global:selectedBranch -eq ‘Branch1’){
        $global:userList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch1"
        $global:groupsList = Get-ADGroup -Filter * -SearchBase "Distinguished Name/Path for groups"
        $global:managerList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch1" -Properties title | Where-Object {$_.title -Like "*Manager*"} | Select name
        $global:deptList = departmentsBranch1
        $global:officeList = @("Location1 Office","Location2 Office")
        $global:companyList = @("Location1 Name","Location 2 Name")
        $global:cityList = @("?","?")
        $global:streetList = @("?","?")
        $global:zipList = @("?","?")
        $global:countryList = @("?","?")
        $global:regionList = @("?","?")
    }
    elseif($global:selectedBranch -eq ’Branch2’){
        $global:userList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch2"
        $global:groupsList = Get-ADGroup -Filter * -SearchBase "Distinguished Name/Path for groups"
        $global:managerList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch2" -Properties title | Where-Object {$_.title -Like "*Manager*"} | Select name
        $global:deptList = departmentsBranch2
        $global:officeList = @("Location1 Office","Location2 Office","Location3 Office")
        $global:companyList = @("Location1 Name","Location 2 Name","Location3 Name")
        $global:cityList = @("?","?","?")
        $global:streetList = @("?","?","?")
        $global:zipList = @("?","?","?")
        $global:countryList = @("?","?","?")
        $global:regionList = @("?","?","?")
    }
    else{
        $global:userList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch3"
        $global:groupsList = Get-ADGroup -Filter * -SearchBase "Distinguished Name/Path for groups"
        $global:managerList = Get-ADUser -Filter * -SearchBase "Distinguished Name/Path Branch3" -Properties title | Where-Object {$_.title -Like "*Manager*"} | Select name
        $global:deptList = departmentsBranch3
        $global:officeList = @("Location1 Office","Location2 Office","Location3 Office","Location4 Office") 
        $global:companyList = @("Location1 Name","Location 2 Name","Location3 Name","Location 4 Name")
        $global:cityList = @("?","?","?","?")
        $global:streetList = @("?","?","?","?")
        $global:zipList = @("?","?","?","?")
        $global:countryList = @("?","?","?","?")
        $global:regionList = @("?","?","?","?")
    }
    
}

<#
Name: userSetup
Purpose: Creating the logon name for the users SAM account Name, user principal name, and emails.It also creates the Entire user.
Parameters: None
Variables: firstNameChar - Separates the first character of the users name so it can be sued to make the auto generated email.
           global:mailOption - Creates the premade email for the new user.
#>
function userSetup(){
    $firstNameChar = $global:userFirstName.ToLower().ToCharArray()
    $global:mailOption = $firstNameChar[0], $global:userLastName.toLower() -join ""

    $flag = 0
    while($flag -eq 0){
        try{
        
            if($global:selectedBranch -eq ‘Branch1’){
                New-ADUser -Name $global:fullName -Path '' -UserPrincipalName $global:primaryMail -DisplayName $global:displayName -GivenName $global:userFirstName -Surname $global:userLastName -Manager $global:selectedManager -Office $global:selectedOffice -Organization $global:selectedOffice -Department $global:selectedDept -Title $global:selectedTitle -Description $global:selectedTitle -Company $global:selectedCompany -StreetAddress $global:selectedStreet -City $global:selectedCity -PostalCode $global:selectedZip -State $global:selectedRegion -Country $global:selectedCountry -SamAccountName $global:mailOption -ChangePasswordAtLogon $False -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText “Password” -Force)
            }
            elseif($global:selectedBranch -eq ’Branch2’){
                New-ADUser -Name $global:fullName -Path '' -UserPrincipalName $global:primaryMail -DisplayName $global:displayName -GivenName $global:userFirstName -Surname $global:userLastName -Manager $global:selectedManager -Office $global:selectedOffice -Organization $global:selectedOffice -Department $global:selectedDept -Title $global:selectedTitle -Description $global:selectedTitle -Company $global:selectedCompany -StreetAddress $global:selectedStreet -City $global:selectedCity -PostalCode $global:selectedZip -State $global:selectedRegion -Country $global:selectedCountry -SamAccountName $global:mailOption -ChangePasswordAtLogon $False -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText “Password” -Force)
            }
            else($global:selectedBranch -eq ’Branch3’){
                New-ADUser -Name $global:fullName -Path '' -UserPrincipalName $global:primaryMail -DisplayName $global:displayName -GivenName $global:userFirstName -Surname $global:userLastName -Manager $global:selectedManager -Office $global:selectedOffice -Organization $global:selectedOffice -Department $global:selectedDept -Title $global:selectedTitle -Description $global:selectedTitle -Company $global:selectedCompany -StreetAddress $global:selectedStreet -City $global:selectedCity -PostalCode $global:selectedZip -State $global:selectedRegion -Country $global:selectedCountry -SamAccountName $global:mailOption -ChangePasswordAtLogon $False -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText “Password” -Force)
            }
			
            $flag = 1
            $distinguished = Get-ADUser -Filter "Name -eq '$global:fullName'" -Properties DistinguishedName
                $global:selectedGroups | ForEach-Object{
                    Add-ADGroupMember -Identity $_ -Members $distinguished
                }
        }#try
        catch{
            emailForm
        }
        
    } #while
}



#Running it all
<#
Name: main
Purpose: Runs all the funtions in propper order and protects the data
Parameters: None
Variables: global:selectedBranch - Check if a branch was selected.
           global:selectedManager - Check if a manager was selected.
           global:selectedOffice - Check if a office was selected.
           global:selectedTitle - Check if a title was selected.
           global:selectedDept - Check if a department was selected.
           global:fullName - Creates the users full name property.
           global:displayName - Creates users display name property.
           firstNameChar - Isolates the first character of the first name from the string and sets it to lowercase so it can be used in the mailOption.
           global:globalMail - Sets up global email.
           global:primaryMail - Sets up primary email.
           global:cancelFlag - Boolean yes or no for weter or not to end the program.
#>
function main(){

    #Location Form
    branchForm
    if($global:cancelFlag -eq 1){ 
       exit   
    }
    while($global:selectedBranch -eq $null){
        branchForm
        if($global:cancelFlag -eq 1){
            exit
        }
    }

    #Properties Forms
    propertiesForm
    if($global:cancelFlag -eq 1){
       exit   
    }
    while(($global:selectedManager -eq $null) -or ($global:selectedOffice -eq $null) -or ($global:selectedTitle -eq "") -or ($global:selectedDept -eq $null)){
        propertiesForm
        if($global:cancelFlag -eq 1){
            exit     
        }
    }

    #Name Form
    nameForm
    if($global:cancelFlag -eq 1){
       exit    
    }
    while(($userFirstName -eq "") -or ($userLastName -eq "")){
        nameForm
        if($global:cancelFlag -eq 1){
            exit    
        }
    }

    $global:fullName = $global:userFirstName, $global:userLastName -Join " "
    $global:displayName = $userLastName, $userFirstName  -Join ", "

    #Groups Form
    groupsForm
    if($global:cancelFlag -eq 1){  
       exit   
    }

    #Make temp Email
    $firstNameChar = $global:userFirstName.ToLower().ToCharArray()
    $global:mailOption = $firstNameChar[0], $global:userLastName.toLower() -join ""

    $global:globalMail = $global:mailOption, "@global.mail.onmicrosoft.com" -join ""
    $global:primaryMail = $global:mailOption, "@email.com" -join ""

    #Make User
    userSetup
    if($global:cancelFlag -eq 1){
        exit
    }
	
	Start-Sleep -Seconds 60
    
    #Email Setup
	Add-PSSNapin Microsoft.Exchange.Management.PowerShell.RecipientManagement
    Enable-RemoteMailbox -Identity $fullName -RemoteRoutingAddress $globalMail
    Set-RemoteMailbox -identity $fullName -EmailAddressPolicyEnabled $false
    Set-RemoteMailbox -identity $fullName -PrimarySmtpAddress $primaryMail

<#
    Write-Host "First Name:" $global:userFirstName
    Write-Host "Last Name:" $global:userLastName
    Write-Host "Full Name:" $global:fullName
    Write-Host "Display Name:" $displayName

    Write-Host "Copied User:" $global:selectedUser
    Write-Host "Manager:" $global:selectedManager
    Write-Host "Office:" $global:selectedOffice
    Write-Host "Department:" $global:selectedDept
    Write-Host "Title:" $global:selectedTitle
    Write-Host "Company:" $global:selectedCompany
    Write-Host "Street:" $global:selectedStreet
    Write-Host "City:" $global:selectedCity 
    Write-Host "Zip:" $global:selectedZip
    Write-Host "Region:" $global:selectedRegion
    Write-Host "Country:" $global:selectedCountry

    Write-Host "Mail Logon:" $global:mailOption
    Write-Host "Global Mail:" $global:globalMail
    Write-Host "Primary Mail:" $global:primaryMail

    Write-Host "Groups:" $global:selectedGroups
#>
}

main

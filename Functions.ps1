function Invoke-DeskproAPI {
    param (
        [string]$Uri
    )
    try {
        $response = Invoke-RestMethod -Uri $Uri -Headers $headers -Method Get
        return $response
    }
    catch {
        $errordata = $_
        switch ($_.Exception.Response.ReasonPhrase) {
            "Not Found" {
                write-host "Ticket does not exist" 
                Break 
            }
            "Unauthorized" { 
                write-host "Skipping Ticket"
                break 
            }
            "Forbidden" { return "forbidden" }
            default {
                Write-Error "API call failed for URI: $Uri"
                Write-Error "Status Code: $($errordata.Exception.Response.StatusCode.value__)"
                Write-Error "Status Description: $($errordata.Exception.Response.StatusDescription)"
                Write-Error "Error Message: $($ErrorData.Exception.Message)"
                throw
            }
        }
    }
}

function get-department ($code) {
    switch ($code) {
        9 { return "UC Infrastructure" }
        21 { return "EBS Service Desk" }
        10 { return "UC Admin" }
    }
}
function createName {
    param ($name)
    
    try {
        $part1 = $name.split("@")[0].split(".")
        $part2 = $part1[0] + " " + $part1[1]
        return $part2
    }
    catch {
        Write-Error "Failed to create name"
        Write-Error "Error: $($_.Exception.Message)"
        if ($null -eq $name) {
            write-error "No value provided"
            return $null
        }
    }
}

function set-uciResolution ($code) {
    switch ($code) {
        # Software
        864 { return 'Software/ONS/New User' }
        865 { return 'Software/ONS/Account Amendment' }
        866 { return 'Software/ONS/Account Removal' }
        867 { return 'Software/ONS/Advice' }
        868 { return 'Software/ONS/Fault' }
                
        882 { return 'Software/Liberty/IVR/Create' }
        883 { return 'Software/Liberty/IVR/Amendment' }
        884 { return 'Software/Liberty/IVR/Removal' }
        870 { return 'Software/Liberty/New Agent' }
        871 { return 'Software/Liberty/New Supervisor' }
        872 { return 'Software/Liberty/Account Amendment' }
        873 { return 'Software/Liberty/Pin Reset' }
        874 { return 'Software/Liberty/Account Removal' }
        885 { return 'Software/Liberty/Voicemail/New' }
        886 { return 'Software/Liberty/Voicemail/Amend' }
        887 { return 'Software/Liberty/Voicemail/Remove' }
        876 { return 'Software/Liberty/Reports' }
        877 { return 'Software/Liberty/Training' }
        878 { return 'Software/Liberty/Recordings' }
        879 { return 'Software/Liberty/Contact Portal' }
        880 { return 'Software/Liberty/Fault' }
        881 { return 'Software/System Config' }
                
        1014 { return 'Software/ASC/Fault' }
        1015 { return 'Software/ASC/New User' }
        1016 { return 'Software/ASC/Maintenance' }
        1017 { return 'Software/ASC/Reports' }
        1018 { return 'Software/ASC/Recording' }
                
        888 { return 'Software/Contact Centre/Concierge' }
        889 { return 'Software/Contact Centre/Manager' }
        988 { return 'Software/Contact Centre/ACWin' }
        1005 { return 'Software/Contact Centre/Oscar' }
                
        890 { return 'Software/Paging/Backups' }
                
        891 { return 'Software/BTS/Reports' }
        892 { return 'Software/BTS/Fault' }
                
        942 { return 'Software/DirX' }
        859 { return 'Software/Jacarta' }
        860 { return 'Software/Server' }
        861 { return 'Software/Recorder' }
        862 { return 'Software/DeskPro' }
        863 { return 'Software/NFF' }
                
        # Hardware
        893 { return 'Hardware/VoIP Phone/Provision' }
        894 { return 'Hardware/VoIP Phone/Replace' }
        895 { return 'Hardware/VoIP Phone/Fault' }
        940 { return 'Hardware/VoIP Phone/Configuration' }
                
        913 { return 'Hardware/Desk Phone/Provision' }
        914 { return 'Hardware/Desk Phone/Replace' }
        939 { return 'Hardware/Desk Phone/Configuration' }
        915 { return 'Hardware/Desk Phone/Move' }
        916 { return 'Hardware/Desk Phone/Office Move' }
        983 { return 'Hardware/Desk Phone/Backup Phones' }
                
        898 { return 'Hardware/Curly Cord' }
                
        917 { return 'Hardware/Intercom/Provision' }
        918 { return 'Hardware/Intercom/Replace' }
        919 { return 'Hardware/Intercom/Fault' }
                
        894 { return 'Hardware/Headset/Provision' }
        895 { return 'Hardware/Headset/Replace' }
                
        920 { return 'Hardware/Paging/Terminal' }
        921 { return 'Hardware/Paging/Recording PCs' }
        922 { return 'Hardware/Paging/Infrastructure' }
        923 { return 'Hardware/Paging/Failover Testing' }
                
        924 { return 'Hardware/DX/Infrastructure' }
        925 { return 'Hardware/4K/Infrastructure' }
        926 { return 'Hardware/OSV/Infrastructure' }
                
        927 { return 'Hardware/DAS/Fault' }
        928 { return 'Hardware/DAS/Signal Check' }
        997 { return 'Hardware/DAS/Provision' }
        998 { return 'Hardware/DAS/Repair' }
                
        929 { return 'Hardware/Server/Commission' }
        930 { return 'Hardware/Server/Decommission' }
        931 { return 'Hardware/Server/Relocate' }
        932 { return 'Hardware/Server/Fault' }
                
        933 { return 'Hardware/UPS/Provision' }
        934 { return 'Hardware/UPS/Replace' }
        935 { return 'Hardware/UPS/Config' }
                
        908 { return 'Hardware/Power' }
                
        936 { return 'Hardware/Aircon/Fault' }
        986 { return 'Hardware/Aircon/Provision' }
        987 { return 'Hardware/Aircon/Replace' }
                
        937 { return 'Hardware/Jacarta/Fault' }
                
        938 { return 'Hardware/Datacentre/Fault' }
                
        1000 { return 'Hardware/Contact Centre PCs' }
                
        1002 { return 'Hardware/Lifts/Configuration' }
        1003 { return 'Hardware/Lifts/New Provide' }
        1004 { return 'Hardware/Lifts/Fault' }
                
        912 { return 'Hardware/NFF' }
                
        # Admin
        842 { return 'Admin/Purchase Order' }
        843 { return 'Admin/Stock Take' }
        844 { return 'Admin/Engineer Assist' }
        845 { return 'Admin/Leavers Report' }
        846 { return 'Admin/Expensive Call Report' }
        847 { return 'Admin/Disposal' }
        848 { return 'Admin/Engineer Access' }
        849 { return 'Admin/Key Request' }
        850 { return 'Admin/Cab Room Maintenance' }
        851 { return 'Admin/Space Request' }
        852 { return 'Admin/Site Provision' }
        853 { return 'Admin/Site Relocation' }
                
        990 { return 'Admin/Advice/Liberty' }
        991 { return 'Admin/Advice/ONS' }
        992 { return 'Admin/Advice/Other' }
                
        default { return 'Unknown Code' }
    }
}

function set-ucaRequest ($code) {
    switch ($code) {
        945 { return 'PO Request' }
        947 { return 'Mobile Phones / New Request' }
        948 { return 'Mobile Phones / Replacement' }
        949 { return 'Dongles' }
        951 { return 'Pagers / New' }
        default { return 'Unknown Code' }
    }
}

function set-EBSSla ($code) {
    switch ($code) {
        647 { return 'INCIDENT / INC P1' }
        648 { return 'INCIDENT / INC P2' }
        649 { return 'INCIDENT / INC P3' }
        650 { return 'INCIDENT / INC P4' }
        652 { return 'PROBLEM / PRB P1' }
        653 { return 'PROBLEM / PRB P2' }
        654 { return 'PROBLEM / PRB P3' }
        655 { return 'PROBLEM / PRB P4' }
        default { return 'Unknown Code' }
    }
}

function Wait-WithProgress {
    param (
        [Parameter(Mandatory=$true)]
        [int]$Seconds,
        
        [Parameter(Mandatory=$false)]
        [string]$Activity = "Halting Script for ",
        
        [Parameter(Mandatory=$false)]
        [string]$Status = "$($seconds) seconds..."
    )
    
    $total = $Seconds * 1000  # Convert seconds to 1000ms intervals for accurate timing
    $current = 0
    
    while ($current -lt $total) {
        $percentComplete = [math]::Round(($current / $total) * 100)
        
        # Update status to show percentage complete
        $statusWithPercent = "$Status ($percentComplete%)"
        
        Write-Progress -Activity $Activity -Status $statusWithPercent -PercentComplete $percentComplete
        
        Start-Sleep -Milliseconds 1000  # Sleep for 1 second
        $current += 1000  # Increment by 1000ms
    }
    
    Write-Progress -Activity $Activity -Completed
}

function Get-PersonName {
    param (
        [int]$personId
    )
    
    try {
        $url = "$baseUrl/people/$personId"
        $person = Invoke-DeskproAPI -Uri $url
        return $person.data.name
    }
    catch {
        Write-Error "Failed to get person information for ID: $personId"
        Write-Error "Error: $($_.Exception.Message)"
        return $personId  # Return the ID if we can't get the name
    }
}
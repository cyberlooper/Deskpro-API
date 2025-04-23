#Script Behavior

$ErrorActionPreference = 'Continue'

#Powershell includes
. ./env.ps1
#. ./global_variables.ps1
. ./Functions.ps1

# Variables

$tickets = @()
$ticketMessages = @()
$csvData = @()

$ticketNumber = 1
$limit = 1000

$totalCalls = 0


## Output file for CSV
$outputFile = ".\Deskpro_to_JIRA_FORCED.csv"

#Functions
## See Functions.ps1

<# ---------------------------------[Main Script]----------------------------------#>
try {
    
    do {
        do {
            #Exceptions now handled in Invoke-DeskproAPI function

            #Get Ticket and Message Data
            $ticketURL = "$baseUrl/tickets/$ticketNumber"
            $ticketMessageURL = "$baseUrl/tickets/$ticketNumber/messages"
            
            #Build ticket data
            Write-Host "Building Data for Ticket $ticketNumber...."
            $ticketresponse = Invoke-DeskproAPI -Uri $ticketURL
            $waitTime = 300
            while (($ticketresponse -eq "Unauthorized" -or ($ticketresponse -eq "Forbidden"))) {
                Write-host "IP Throttling hit. Waiting to appease API"
                Wait-WithProgress -Seconds $waitTime
                $waitTime += 300

                $messageResponse = Invoke-DeskproAPI -Uri $ticketMessageURL
            }
            $totalCalls++

            #start-sleep -seconds 1
            if (!($null -eq $ticketresponse)) {
                $tickets += $ticketresponse.data
            }
            
            #Build Message Data
            $messageResponse = Invoke-DeskproAPI -Uri $ticketMessageURL 
            $waitTime = 300
            while (($messageResponse -eq "Unauthorized" -or ($ticketresponse -eq "Forbidden"))) {
                Write-host "IP Throttling hit. Waiting to appease API"
                Wait-WithProgress -Seconds $waitTime
                $waitTime = $waitTime + 300

                $messageResponse = Invoke-DeskproAPI -Uri $ticketMessageURL
            }
            $totalCalls++
            
            #start-sleep -seconds 1
            if (!($null -eq $messageResponse)) {
                $ticketMessages += $messageResponse.data
            }
            $totalCalls++

            $ticketNumber++

            start-sleep -seconds 1

        } until ($ticketNumber -eq $limit)

        $totalTickets += $limit
        $limit += 5000
        
        #Process data into $csvData
        
        foreach ($ticket in $tickets) {
            $id = $ticket.id
            $tmpmessages = $ticketMessages.Where({ $_.ticket -eq $id })
            
            #Clear old messages ready for new ticket
            $messageFields = @()
            foreach ($message in $tmpmessages) {              
                # Strip HTML using regex
                $cleanMessage = [regex]::Replace($message.message, '<.*?>', '')
                # Add to message fields array
                $messageFields += "From: $($message.person) - $($message.date_created) : $($cleanMessage -replace "`r?`n", ' ')"
            }
            
            # Add ticket data to CSV
            $csvRow = [PSCustomObject]@{
                "TicketID"               = $ticket.id
                "UniqueID"               = $ticket.ref
                "Team"                   = $ticket.department
                "Summary"                = $ticket.subject
                "Description"            = $ticket.content
                
                # Shared fields
                "MethodRaised"           = $ticket.fields.'170'.detail.PSObject.Properties.value.title
                "Supplier Ref"           = $ticket.fields.'32'.detail.PSObject.Properties.value
                "Location & HR Number"   = $ticket.fields.'42'.detail.PSObject.Properties.value
                "Contact Number"         = $ticket.fields.'43'.detail.PSObject.Properties.value
        
                # EBS Fields
                "EBSTicketType"          = $ticket.fields.'81'.detail.PSObject.Properties.value.title
                "EBSResolutionType"      = $ticket.fields.'186'.detail.PSObject.Properties.value.title
                "EBSSLA"                 = set-ebssla -code $ticket.fields.'645'.detail.PSObject.Properties.value.id
                "TrustIT"                = $ticket.fields.'175'.detail.PSObject.Properties.value.title
        
                # UC Infrastructure Fields
                "UCITicketType"          = $ticket.fields.'194'.detail.PSObject.Properties.value.title
                "UCIResolutionType"      = set-uciresolution -code $ticket.fields.'167'.detail.PSObject.Properties.value.ID
                "UCSLA"                  = $ticket.fields.'624'.detail.PSObject.Properties.value.title
                
                # UC Admin Fields
                "UCABudgetHolder"        = $ticket.fields.'153'.detail.PSObject.Properties.value
                "UCASparePagerNumber"    = $ticket.fields.'151'.detail.PSObject.Properties.value
                
                ##Conditional Fields
                "UCARequest"             = set-ucaRequest -code $ticket.fields.'944'.detail.PSObject.Properties.value.ID
                "UCADivision"            = $ticket.fields.'802'.detail.PSObject.Properties.value.title
                "UCACost"                = $ticket.fields.'956'.detail.PSObject.Properties.value
                "UCAccountCode"          = $ticket.fields.'957'.detail.PSObject.Properties.value.title
                "UCAFunding"             = $ticket.fields.'971'.detail.PSObject.Properties.value.title
                "UCARequiredDescription" = $ticket.fields.'974'.detail.PSObject.Properties.value.title
                "UCAExceed10k"           = $ticket.fields.'975'.detail.PSObject.Properties.value.title
                "UCAProcProcFollowed"    = $ticket.fields.'978'.detail.PSObject.Properties.value.title
                "UCAProcProcVal"         = $ticket.fields.'981'.detail.PSObject.Properties.value.title
        
                #Footer data
                "TicketStatus"           = $ticket.ticket_status
                "Reporter"               = createName $ticket.person_email
                "ReporterEmail"          = $ticket.person_email
                "Assignee"               = if ($ticket.agent) { $ticket.agent.name } else { "Unassigned" }
                "Created"                = $ticket.date_created
                "Updated"                = $ticket.date_updated
                "Comments"               = $commentHistory
        
                #MessageData
                "Message1"               = if ($messageFields.Count -ge 1) { $messageFields[0] } else { $null }
                "Message2"               = if ($messageFields.Count -ge 2) { $messageFields[1] } else { $null }
                "Message3"               = if ($messageFields.Count -ge 3) { $messageFields[2] } else { $null }
                "Message4"               = if ($messageFields.Count -ge 4) { $messageFields[3] } else { $null }
                "Message5"               = if ($messageFields.Count -ge 5) { $messageFields[4] } else { $null }
                "Message6"               = if ($messageFields.Count -ge 6) { $messageFields[5] } else { $null }
                "Message7"               = if ($messageFields.Count -ge 7) { $messageFields[6] } else { $null }
                "Message8"               = if ($messageFields.Count -ge 8) { $messageFields[7] } else { $null }
                "Message9"               = if ($messageFields.Count -ge 9) { $messageFields[8] } else { $null }
                "Message10"              = if ($messageFields.Count -ge 10) { $messageFields[9] } else { $null }
                "Message11"              = if ($messageFields.Count -ge 11) { $messageFields[10] } else { $null }
                "Message12"              = if ($messageFields.Count -ge 12) { $messageFields[11] } else { $null }
                "Message13"              = if ($messageFields.Count -ge 13) { $messageFields[12] } else { $null }
                "Message14"              = if ($messageFields.Count -ge 14) { $messageFields[13] } else { $null }
                "Message15"              = if ($messageFields.Count -ge 15) { $messageFields[14] } else { $null }
                "Message16"              = if ($messageFields.Count -ge 16) { $messageFields[15] } else { $null }
                "Message17"              = if ($messageFields.Count -ge 17) { $messageFields[16] } else { $null }
                "Message18"              = if ($messageFields.Count -ge 18) { $messageFields[17] } else { $null }
                "Message19"              = if ($messageFields.Count -ge 19) { $messageFields[18] } else { $null }
                "Message20"              = if ($messageFields.Count -ge 20) { $messageFields[19] } else { $null }
                "Message21"              = if ($messageFields.Count -ge 21) { $messageFields[20] } else { $null }
                "Message22"              = if ($messageFields.Count -ge 22) { $messageFields[21] } else { $null }
                "Message23"              = if ($messageFields.Count -ge 23) { $messageFields[22] } else { $null }
                "Message24"              = if ($messageFields.Count -ge 24) { $messageFields[23] } else { $null }
                "Message25"              = if ($messageFields.Count -ge 25) { $messageFields[24] } else { $null }
                "Message26"              = if ($messageFields.Count -ge 26) { $messageFields[25] } else { $null }
                "Message27"              = if ($messageFields.Count -ge 27) { $messageFields[26] } else { $null }
                "Message28"              = if ($messageFields.Count -ge 28) { $messageFields[27] } else { $null }
                "Message29"              = if ($messageFields.Count -ge 29) { $messageFields[28] } else { $null }
                "Message30"              = if ($messageFields.Count -ge 30) { $messageFields[29] } else { $null }
                "Message31"              = if ($messageFields.Count -ge 31) { $messageFields[30] } else { $null }
                "Message32"              = if ($messageFields.Count -ge 32) { $messageFields[31] } else { $null }
                "Message33"              = if ($messageFields.Count -ge 33) { $messageFields[32] } else { $null }
                "Message34"              = if ($messageFields.Count -ge 34) { $messageFields[33] } else { $null }
                "Message35"              = if ($messageFields.Count -ge 35) { $messageFields[34] } else { $null }
                "Message36"              = if ($messageFields.Count -ge 36) { $messageFields[35] } else { $null }
                "Message37"              = if ($messageFields.Count -ge 37) { $messageFields[36] } else { $null }
                "Message38"              = if ($messageFields.Count -ge 38) { $messageFields[37] } else { $null }
                "Message39"              = if ($messageFields.Count -ge 39) { $messageFields[38] } else { $null }
                "Message40"              = if ($messageFields.Count -ge 40) { $messageFields[39] } else { $null }
                "Message41"              = if ($messageFields.Count -ge 41) { $messageFields[40] } else { $null }
                "Message42"              = if ($messageFields.Count -ge 42) { $messageFields[41] } else { $null }
                "Message43"              = if ($messageFields.Count -ge 43) { $messageFields[42] } else { $null }
                "Message44"              = if ($messageFields.Count -ge 44) { $messageFields[43] } else { $null }
                "Message45"              = if ($messageFields.Count -ge 45) { $messageFields[44] } else { $null }
                "Message46"              = if ($messageFields.Count -ge 46) { $messageFields[45] } else { $null }
                "Message47"              = if ($messageFields.Count -ge 47) { $messageFields[46] } else { $null }
                "Message48"              = if ($messageFields.Count -ge 48) { $messageFields[47] } else { $null }
                "Message49"              = if ($messageFields.Count -ge 49) { $messageFields[48] } else { $null }
                "Message50"              = if ($messageFields.Count -ge 50) { $messageFields[49] } else { $null }
            }
            write-host "Created row for DP$($ticket.id) / $($ticket.ref)"
            $csvData += $csvRow
        }

        #clear variables ready for next batch

        $tickets = @()
        $ticketMessages = @()
        $messageFields = @()

        Wait-WithProgress -seconds 300
    } until ($ticketNumber -eq 75000)

    $csvData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8 -Delimiter "^"

}
catch {
    <#Do this if a terminating exception happens#>
}
﻿<#
.SYNOPSIS
Provides SCSM Exchange Connector functionality through PowerShell

.DESCRIPTION
This PowerShell script/runbook aims to address shortcomings and wants in the
out of box SCSM Exchange Connector as well as help enable new functionality around
Work Item scenarios, runbook automation, 3rd party customizations, and
enabling other organizational level processes via email

.NOTES
Author: Adam Dzyacky
Contributors: Martin Blomgren, Leigh Kilday
Reviewers: Tom Hendricks, Brian Weist
Inspiration: The Cireson Community, Anders Asp, Stefan Roth, and (of course) Travis Wright for SMlets examples
Requires: PowerShell 4+, SMlets, and Exchange Web Services API (already installed on SCSM workflow server).
    3rd party option: If you're a Cireson customer and make use of their paid SCSM Portal with HTML Knowledge Base this will work as is
        if you aren't, you'll need to create your own Type Projection for Change Requests for the Add-ChangeRequestComment
        function. Navigate to that function to read more. If you don't make use of their HTML KB, you'll want to keep $searchCiresonHTMLKB = $false
    Signged/Encrypted option: .NET 4.5 is required to use MimeKit.dll
Misc: The Release Record functionality does not exist in this as no out of box (or 3rd party) Type Projection exists to serve this purpose.
    You would have to create your own Type Projection in order to leverage this.
Version: 1.3 = created Set-CiresonPortalAnnouncement and Set-CoreSCSMAnnouncement to introduce announcement integration into the connector
                    by leveraging the new configurable [announcement] keyword and #low or #high tags to set priority on the announcement.
                    absence of the tag results in a normal priority announcement being created
                created Get-SCSMAuthorizedAnnouncer to verify the sender's permissions to post announcements
                created Get-CiresonPortalAnnouncements to search/update announcements
                created Read-MIMEMessage to allow parsing digitally signed or encryped emails. This feature leverages the
                    open source project known as MimeKit by Jeffrey Stedfast. It can be found here - https://github.com/jstedfast/MimeKit   
                created Get-CiresonPortalUser to query for a user through the Cireson Web API to retrieve user information (object)
                created Get-CiresonPortalGroup to query for a group through the Cireson Web API to retrieve group information (object)                          
                created Search-AvailableCiresonPortalOfferings in order to look for relevant request offerings within a user's
                    Service Catalog scope to suggest relevant requests to the Affected User based on the content of their email           
                improved/simplified Search-CiresonKnowledgeBase by use of new Get-CiresonPortalUser function              
                created Get-SCOMDistributedAppHealth (SCOM integration) allows an authorized user to retrieve the health of
                    a distributed application from Operations Manager. Features configurable [keyword].       
                created Get-SCOMAuthorizedRequester (SCOM integration) to verify that the individual requesting status on a SCOM Distributed Application
                    is authorized to do so
Version: 1.2 = created Send-EmailFromWorkflowAccount for future functions to leverage the SCSM workflow account defined therein
                updated Search-CiresonKnowledgeBase to use Send-EmailFromWorkflowAccount
                created $exchangeAuthenticationType so as to introduce Windows Authentication or Impersonation to bring to closer parity with stock EC connector
                expanded email processing loop to prepare for things other than IPM.Note message class (i.e. Calendar appointments, custom message classes per org.)
                created Schedule-WorkItem function to enable setting Scheduled Start/End Dates on Work Items based on the Calendar Start/End times.
                    introduced configuration variable for this feature ($processCalendarAppointment)
                updated Attach-EmailToWorkItem so the Exchange Conversation ID is written into the Description as "ExchangeConversationID:$id;"
                issue on Attach-EmailToWorkItem/Attach-FileToWorkItem where the "AttachedBy" relationship was using the wrong variable
                created Verify-WorkItem to attempt to begin identifying potentially quickly responded messages/append them to
                    current/recently created Work Items. Trying to address scenario where someone emails WF and CC's others. If the others
                    reply back before a notification about the current Work Item ID goes out (or they ignore it), they queue more messages
                    for the connector and in turn create more default work items rather than updating the "original" thread/Work Item that
                    was created in the same/previous processing loop. Also looking to use this function to potentially address 3rd party
                    ticketing systems.
                    Introduced configuration variable for this feature ($mergeReplies)
Version: 1.1 = GitHub issue raised on updating work items. Per discussion was pinpointed to the
                Get-WorkItem function wherein passed in values were including brackets in the search (i.e. [IRxxxx] instead of IRxxxx). Also
                updated the email subject matching regex, so that the Update-WorkItem took the $result.id instead of the $matches[0]. Again, this
                ensures the brackets aren't passed when performing the search/update.
#>

#region #### Configuration ####
#define the an SCSM management server, this could be a remote name or localhost
$scsmMGMTServer = ""

#define/use SCSM WF credentials
#$exchangeAuthenticationType - "windows" or "impersonation" are valid inputs here.
    #Windows will use the credentials that start this script in order to authenticate to Exchange and retrieve messages
        #choosing this option only requires the $workflowEmailAddress variable to be defined
        #this is ideal if you'll be using Task Scheduler or SMA to initiate this
    #Impersonation will use the credentials that are defined here to connect to Exchange and retrieve messages
        #choosing this option requires the $workflowEmailAddress, $username, $password, and $domain variables to be defined
$exchangeAuthenticationType = "windows"
$workflowEmailAddress = ""
$username = ""
$password = ""
$domain = ""

#defaultNewWorkItem = set to either "ir" or "sr"
#minFileSizeInKB = Set the minimum file size in kilobytes to be attached to work items
#createUsersNotInCMDB = If someone from outside your org emails into SCSM this allows you to take that email and create a User in your CMDB
#includeWholeEmail = If long chains get forwarded into SCSM, you can choose to write the whole email to a single action log entry OR the beginning to the first finding of "From:"
#attachEmailToWorkItem = If $true, attach email as an *.eml to each work item. Additionally, write the Exchange Conversation ID into the Description of the Attachment object
#fromKeyword = If $includeWholeEmail is set to true, messages will be parsed UNTIL they find this word
$defaultNewWorkItem = "ir"
$defaultIRTemplate = Get-SCSMObjectTemplate -DisplayName "IR Template Name Goes Here" -computername $scsmMGMTServer
$defaultSRTemplate = Get-SCSMObjectTemplate -DisplayName "SR Template Name Goes Here" -computername $scsmMGMTServer
$minFileSizeInKB = "25"
$createUsersNotInCMDB = $true
$includeWholeEmail = $false
$attachEmailToWorkItem = $false
$fromKeyword = "From"

#processCalendarAppointment = If $true, scheduling appointments with the Workflow Inbox where a [WorkItemID] is in the Subject will
    #set the Scheduled Start and End Dates on the Work Item per the Start/End Times of the calendar appointment
    #and will also override $attachEmailToWorkItem to be $true if set to $false
#processDigitallySignedMessages = If $true, MimeKit will parse digitally signed email messages will also be processed in accordance with
    #settings defined for normal email processing
#processEncryptedMessage = If $true, MimeKit will parse encrypted email messages in accordance with settings defined for normal email processing.
    #In order for this to work, the correct decrypting certificate must be placed in either the Current User or Local Machine store
#certStore = If you will be processing encrypted email, you must define where the decrypting certificate is located. This takes the values
    #of either "user" or "machine"
#mergeReplies = If $true, emails that are Replies (signified by RE: in the subject) will attempt to be matched to a Work Item in SCSM by their
    #Exchange Conversation ID and will also override $attachEmailToWorkItem to be $true if set to $false
$processCalendarAppointment = $false
$processDigitallySignedMessages = $false
$processEncryptedMessages = $false
$certStore = "user"
$mergeReplies = $false

#optional, enable integration with Cireson Knowledge Base/Service Catalog
#this uses the now depricated Cireson KB API Search by Text, it works as of v7.x but should be noted it could be entirely removed in future portals
#$numberOfWordsToMatchFromEmailToRO = defines the minimum number of words that must be matched from an email/new work item before Knowledge Articles will be
    #suggested to the Affected User about them
#searchAvailableCiresonPortalOfferings = search available Request Offerings within the Affected User's permission scope based words matched in
    #their email/new work item
#$ciresonPortalServer = URL that will be used to search for KB articles via invoke-webrequest. Make sure to leave the "/" after your tld!
#$ciresonPortalWindowsAuth = how invoke-webrequest should attempt to authenticate to your portal server.
    #Leave true if your portal uses Windows Auth, change to False for Forms authentication.
    #If using forms, you'll need to set the ciresonPortalUsername and Password variables. For ease, you could set this equal to the username/password defined above
$searchCiresonHTMLKB = $false
$numberOfWordsToMatchFromEmailToRO = 1
$searchAvailableCiresonPortalOfferings = $false
$ciresonPortalServer = "https://portalserver.domain.tld/"
$ciresonPortalWindowsAuth = $true
$ciresonPortalUsername = ""
$ciresonPortalPassword = ""

#optional, enable Announcement control in SCSM/Cireson portal from email
#enableSCSMAnnouncements/enableCiresonPortalAnnouncements: You can create/update announcements
    #in Core SCSM or the Cireson Portal by changing these values from $false to $true
#announcementKeyword = if this [keyword] is in the message body, a new announcement will be created
#approved users, groups, type = control who is authorized to post announcements to SCSM/Cireson Portal
    #you can configure individual users by email address or use an Active Directory group
#priority keywords: These are the words that you can also include in the body of your message to further
    #define an announcment by setting it's priority.  For example the body of your message could be "Patching systems this weekend. [announcement] #low"
#priorityExpirationInHours: Since both SCSM and the Cireson require an announcement expiration date, when announcements are created
    #this is the number of hours added to the current time to set the announcement to expire. If you send Calendar Meetings which by definition
    #have a start and end time, these expirationInHours style variables are ignored
$enableSCSMAnnouncements = $false
$enableCiresonPortalAnnouncements = $false
$announcementKeyword = "announcement"
$approvedADGroupForSCSMAnnouncements = "my custom AD SCSM Authorized Announcers Users group"
$approvedUsersForSCSMAnnouncements = "myfirst.email@domain.com", "mysecond.address@domain.com"
$approvedMemberTypeForSCSMAnnouncer = "group"
$lowAnnouncemnentPriorityKeyword = "low"
$criticalAnnouncemnentPriorityKeyword = "high"
$lowAnnouncemnentExpirationInHours = 7
$normalAnnouncemnentExpirationInHours = 3
$criticalAnnouncemnentExpirationInHours = 1

#optional, enable SCOM functionality
#enableSCOMIntegration = set to $true or $false to enable this functionality
#scomMGMTServer = set equal to the name of your scom management server
#approvedMemberTypeForSCOM = To prevent unapproved individuals from gaining knowledge about your SCOM environment via SCSM, you must
    #choose to set either an AD Group that the sender must be a part of or you must manually define email addresses of users allowed to
    #make these requests. This variable can be set to "users" or "group"
#approvedADGroupForSCOM = if approvedUsersForSCOM = group, set this to the AD Group that contains groups/members that are allowed to make SCOM email requests
    #this approach allows you control access through Active Directory
#approvedUsersForSCOM = if approvedUsersForSCOM = users, set this to a comma seperated list of email addresses that are allowed to make SCOM email requests
    #this approach allows you to control through this script
#distributedApplicationHealthKeyword = the keyword to use in the subject for the connector to request DA status from SCOM
$enableSCOMIntegration = $false
$scomMGMTServer = ""
$approvedMemberTypeForSCOM = "group"
$approvedADGroupForSCOM = "my custom AD SCOM Authorized Users group"
$approvedUsersForSCOM = "myfirst.email@domain.com", "mysecond.address@domain.com"
$distributedApplicationHealthKeyword = "health"

#define SCSM Work Item keywords to be used
$acknowledgedKeyword = "acknowledge"
$reactivateKeyword = "reactivate"
$resolvedKeyword = "resolved"
$closedKeyword = "closed"
$holdKeyword = "hold"
$cancelledKeyword = "cancelled"
$takeKeyword = "take"
$completedKeyword = "completed"
$skipKeyword = "skipped"
$approvedKeyword = "approved"
$rejectedKeyword = "rejected"

#define the path to the Exchange Web Services API and MimeKit
$exchangeEWSAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
$mimeKitDLLPath = "c:\smletsExchangeConnector\mimekit.dll"

#enable logging per standard Exchange Connector registry keys
#valid options on that registry key are 1 to 7 where 7 is the most verbose
#$loggingLevel = (Get-ItemProperty "HKLM:\Software\Microsoft\System Center Service Manager Exchange Connector" -ErrorAction SilentlyContinue).LoggingLevel
#$loggingLevel = 1

#endregion

#region #### SCSM Classes ####
$irClass = get-scsmclass "System.WorkItem.Incident$" -computername $scsmMGMTServer
$srClass = get-scsmclass "System.WorkItem.ServiceRequest$" -computername $scsmMGMTServer
$prClass = get-scsmclass "System.WorkItem.Problem$" -computername $scsmMGMTServer
$crClass = get-scsmclass "System.Workitem.ChangeRequest$" -computername $scsmMGMTServer
$rrClass = get-scsmclass "System.Workitem.ReleaseRecord$" -computername $scsmMGMTServer
$maClass = get-scsmclass "System.WorkItem.Activity.ManualActivity$" -computername $scsmMGMTServer
$raClass = get-scsmclass "System.WorkItem.Activity.ReviewActivity$" -computername $scsmMGMTServer
$paClass = get-scsmclass "System.WorkItem.Activity.ParallelActivity$" -computername $scsmMGMTServer
$saClass = get-scsmclass "System.WorkItem.Activity.SequentialActivity$" -computername $scsmMGMTServer
$daClass = get-scsmclass "System.WorkItem.Activity.DependentActivity$" -computername $scsmMGMTServer

$raHasReviewerRelClass = Get-SCSMRelationshipClass "System.ReviewActivityHasReviewer$" -computername $scsmMGMTServer
$raReviewerIsUserRelClass = Get-SCSMRelationshipClass "System.ReviewerIsUser$" -computername $scsmMGMTServer
$raVotedByUserRelClass = Get-SCSMRelationshipClass "System.ReviewerVotedByUser$" -computername $scsmMGMTServer

$userClass = get-scsmclass "System.User$" -computername $scsmMGMTServer
$domainUserClass = get-scsmclass "System.Domain.User$" -computername $scsmMGMTServer
$notificationClass = get-scsmclass "System.Notification.Endpoint$" -computername $scsmMGMTServer

$irLowImpact = Get-SCSMEnumeration "System.WorkItem.TroubleTicket.ImpactEnum.Low$" -computername $scsmMGMTServer
$irLowUrgency = Get-SCSMEnumeration "System.WorkItem.TroubleTicket.ImpactEnum.Low$" -computername $scsmMGMTServer
$irActiveStatus = Get-SCSMEnumeration "IncidentStatusEnum.Active$" -computername $scsmMGMTServer

$affectedUserRelClass = get-scsmrelationshipclass "System.WorkItemAffectedUser$" -computername $scsmMGMTServer
$assignedToUserRelClass  = Get-SCSMRelationshipClass "System.WorkItemAssignedToUser$" -computername $scsmMGMTServer
$createdByUserRelClass = Get-SCSMRelationshipClass "System.WorkItemCreatedByUser$" -computername $scsmMGMTServer
$workResolvedByUserRelClass = Get-SCSMRelationshipClass "System.WorkItem.TroubleTicketResolvedByUser$" -computername $scsmMGMTServer
$wiRelatesToCIRelClass = Get-SCSMRelationshipClass "System.WorkItemRelatesToConfigItem$" -computername $scsmMGMTServer
$wiRelatesToWIRelClass = Get-SCSMRelationshipClass "System.WorkItemRelatesToWorkItem$" -computername $scsmMGMTServer
$wiContainsActivityRelClass = Get-SCSMRelationshipClass "System.WorkItemContainsActivity$" -computername $scsmMGMTServer
$sysUserHasPrefRelClass = Get-SCSMRelationshipClass "System.UserHasPreference$" -ComputerName $scsmMGMTServer

$fileAttachmentClass = Get-SCSMClass -Name "System.FileAttachment$" -computername $scsmMGMTServer
$fileAttachmentRelClass = Get-SCSMRelationshipClass "System.WorkItemHasFileAttachment$" -computername $scsmMGMTServer
$fileAddedByUserRelClass = Get-SCSMRelationshipClass "System.FileAttachmentAddedByUser$" -ComputerName $scsmMGMTServer
$managementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $scsmMGMTServer

$irTypeProjection = Get-SCSMTypeProjection "system.workitem.incident.projectiontype$" -computername $scsmMGMTServer
$srTypeProjection = Get-SCSMTypeProjection "system.workitem.servicerequestprojection$" -computername $scsmMGMTServer

$userHasPrefProjection = Get-SCSMTypeProjection "System.User.Preferences.Projection$" -computername $scsmMGMTServer
#endregion

#region #### Exchange Connector Functions ####
function New-WorkItem ($message, $wiType, $returnWIBool) 
{
    $from = $message.From
    $to = $message.To
    $cced = $message.CC
    $title = $message.subject
    $description = $message.body

    #if the message is longer than 4000 characters take only the first 4000.
    if ($description.length -ge "4000")
    {
        $description = $description.substring(0,4000)
    }

    #find Affected User from the From Address
    $relatedUsers = @()
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$from'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
    if ($userSMTPNotification) 
    { 
        $affectedUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $affectedUser = create-userincmdb $from
        }
    }

    #find Related Users (To)       
    if ($to.count -gt 0)
    {
        if ($to.count -eq 1)
        {
            $userToSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($to.address)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
            if ($userToSMTPNotification) 
            { 
                $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userToSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $relatedUser = create-userincmdb $to.address
                    $relatedUsers += $relatedUser
                }
            }
        }
        else
        {
            $x = 0
            while ($x -lt $to.count)
            {
                $ToSMTP = $to[$x]
                $userToSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($ToSMTP.address)'"  -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
                if ($userToSMTPNotification) 
                { 
                    $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userToSMTPNotification -computername $scsmMGMTServer).sourceObject 
                    $relatedUsers += $relatedUser
                }
                else
                {
                    if ($createUsersNotInCMDB -eq $true)
                    {
                        $relatedUser = create-userincmdb $ToSMTP.address
                        $relatedUsers += $relatedUser
                    }
                }
                $x++
            }
        }
    }
    
    #find Related Users (Cc)         
    if ($cced.count -gt 0)
    {
        if ($cced.count -eq 1)
        {
            $userCCSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($cced.address)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
            if ($userCCSMTPNotification) 
            { 
                $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userCCSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $relatedUser = create-userincmdb $cced.address
                    $relatedUsers += $relatedUser
                }
            }
        }
        else
        {
            $x = 0
            while ($x -lt $cced.count)
            {
                $ccSMTP = $cced[$x]
                $userCCSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($ccSMTP.address)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
                if ($userCCSMTPNotification) 
                { 
                    $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userCCSMTPNotification -computername $scsmMGMTServer).sourceObject 
                    $relatedUsers += $relatedUser
                }
                else
                {
                    if ($createUsersNotInCMDB -eq $true)
                    {
                        $relatedUser = create-userincmdb $ccSMTP.address
                        $relatedUsers += $relatedUser
                    }
                }
                $x++
            }
        }
    }

    #create the Work Item based on the globally defined Work Item type and Template
    switch ($defaultNewWorkItem) 
    {
        "ir" {
                    $newWorkItem = New-SCSMObject -Class $irClass -PropertyHashtable @{"ID" = "IR{0}"; "Status" = $irActiveStatus; "Title" = $title; "Description" = $description; "Classification" = $null; "Impact" = $irLowImpact; "Urgency" = $irLowUrgency; "Source" = "IncidentSourceEnum.Email$"} -PassThru -computername $scsmMGMTServer
                    $irProjection = Get-SCSMObjectProjection -ProjectionName $irTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -Projection $irProjection -Template $defaultIRTemplate -computername $scsmMGMTServer
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk -computername $scsmMGMTServer
                        }
                    }
                
                    #### Determine auto-response logic for Knowledge Base and/or Request Offering Search ####
                    if (($searchCiresonHTMLKB -eq $true) -and ($searchAvailableCiresonPortalOfferings -eq $true))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Knowledge Base Articles and matching Request Offering URLs
                        $kbURLs = Search-CiresonKnowledgeBase -workItem $newWorkItem -ciresonPortalUser $portalUser
                        $requestURLs = Search-AvailableCiresonPortalOfferings -ciresonPortalUser $portalUser -workItem $newWorkItem

                        #combine KB results and Offering results into a single email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$resolvedKeyword]" + "`">resolve</a>"
                        $emailBodyResponse = "We found some knowledge articles and requests that may be of assistance to you <br/><br/>
                        Knowledge Articles: <br/><br />
                        $kbURLs<br /><br />
                        Requests: <br /><br />
                        $requestURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    elseif (($searchCiresonHTMLKB -eq $true) -and ($searchAvailableCiresonPortalOfferings -eq $false))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Knowledge Base Articles URLs
                        $kbURLs = Search-CiresonKnowledgeBase -workItem $newWorkItem -ciresonPortalUser $portalUser

                        #prepare KB result email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$resolvedKeyword]" + "`">resolve</a>"
                        $emailBodyResponse = "We found some knowledge articles that may be of assistance to you <br/><br/>
                        Knowledge Articles: <br/><br />
                        $kbURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    elseif (($searchCiresonHTMLKB -eq $false) -and ($searchAvailableCiresonPortalOfferings -eq $true))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Request Offering URLs
                        $requestURLs = Search-AvailableCiresonPortalOfferings -ciresonPortalUser $portalUser -workItem $newWorkItem

                        #prepare Request Offering results email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$resolvedKeyword]" + "`">resolve</a>"
                        $emailBodyResponse = "We found some requests on the portal that help you get what you need faster <br/><br/>
                        Knowledge Articles: <br/><br />
                        $requestURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    else
                    {
                        #both options are set to $false
                        #don't suggest anything to the Affected User based on their recently created Default Work Item
                    }
                }
        "sr" {
                    $newWorkItem = new-scsmobject -class $srClass -propertyhashtable @{"ID" = "SR{0}"; "Title" = $title; "Description" = $description; "Status" = "ServiceRequestStatusEnum.Submitted$"} -PassThru -computername $scsmMGMTServer
                    $srProjection = Get-SCSMObjectProjection -ProjectionName $srTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -projection $srProjection -Template $defaultSRTemplate -computername $scsmMGMTServer
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk -computername $scsmMGMTServer
                        }
                    }
                    
                    #### Determine auto-response logic for Knowledge Base and/or Request Offering Search ####
                    if (($searchCiresonHTMLKB -eq $true) -and ($searchAvailableCiresonPortalOfferings -eq $true))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Knowledge Base Articles and matching Request Offering URLs
                        $kbURLs = Search-CiresonKnowledgeBase -workItem $newWorkItem -ciresonPortalUser $portalUser
                        $requestURLs = Search-AvailableCiresonPortalOfferings -ciresonPortalUser $portalUser -workItem $newWorkItem

                        #combine KB results and Offering results into a single email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$cancelledKeyword]" + "`">cancel</a>"
                        $emailBodyResponse = "We found some knowledge articles and requests that may be of assistance to you <br/><br/>
                        Knowledge Articles: <br/><br />
                        $kbURLs<br /><br />
                        Requests: <br /><br />
                        $requestURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    elseif (($searchCiresonHTMLKB -eq $true) -and ($searchAvailableCiresonPortalOfferings -eq $false))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Knowledge Base Articles URLs
                        $kbURLs = Search-CiresonKnowledgeBase -workItem $newWorkItem -ciresonPortalUser $portalUser

                        #prepare KB result email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$cancelledKeyword]" + "`">cancel</a>"
                        $emailBodyResponse = "We found some knowledge articles that may be of assistance to you <br/><br/>
                        Knowledge Articles: <br/><br />
                        $kbURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    elseif (($searchCiresonHTMLKB -eq $false) -and ($searchAvailableCiresonPortalOfferings -eq $true))
                    {
                        #get the user object from the Cireson Portal
                        $portalUser = Get-CiresonPortalUser -username $affectedUser.UserName -domain $affectedUser.Domain

                        #get matching Request Offering URLs
                        $requestURLs = Search-AvailableCiresonPortalOfferings -ciresonPortalUser $portalUser -workItem $newWorkItem

                        #prepare Request Offering results email back to the Affected User
                        $resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $newWorkItem.id + "]" + "&body=This%20can%20be%20[$cancelledKeyword]" + "`">cancel</a>"
                        $emailBodyResponse = "We found some requests on the portal that help you get what you need faster <br/><br/>
                        Knowledge Articles: <br/><br />
                        $requestURLs<br /><br />
                        If any of the above helped you out, you can $resolveMailTo your original request."
                        
                        #send the message
                        Send-EmailFromWorkflowAccount -subject "[$($newWorkItem.id)] - $($newWorkItem.title)" -body $emailBodyResponse -bodyType "HTML" -toRecipients $from
                    }
                    else
                    {
                        #both options are set to $false
                        #don't suggest anything to the Affected User based on their recently created Default Work Item
                    }
                } 
    }

    if ($returnWIBool -eq $true)
    {
        return $newWorkItem
    }
}

function Update-WorkItem ($message, $wiType, $workItemID) 
{
    #determine the comment to add and ensure it's less than 4000 characters
    if ($includeWholeEmail -eq $true)
    {
        $commentToAdd = $message.body
        if ($commentToAdd.length -ge "4000")
        {
            $commentToAdd.substring(0, 4000)
        }
    }
    else
    {
        $fromKeywordPosition = $message.Body.IndexOf("$fromKeyword" + ":")
        if (($fromKeywordPosition -eq $null) -or ($fromKeywordPosition -eq -1))
        {
            $commentToAdd = $message.body
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd.substring(0, 4000)
            }
        }
        else
        {
            $commentToAdd = $message.Body.substring(0, $fromKeywordPosition)
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd.substring(0, 4000)
            }
        }
    }

    #determine who left the comment
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.From)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
    if ($userSMTPNotification) 
    { 
        $commentLeftBy = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $commentLeftBy = create-userincmdb $message.From
        }
    }

    #add any attachments
    if ($message.Attachments)
    {
        Attach-FileToWorkItem $message $workItemID
    }

    #update the work item with the comment and/or action
    switch ($wiType) 
    {
        #### primary work item types ####
        "ir" {
                    $workItem = get-scsmobject -class $irClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                        $assignedToSMTP.TargetAddress {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                        default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                        "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() -computername $scsmMGMTServer}}
                        "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Resolved$" -computername $scsmMGMTServer; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Closed$" -computername $scsmMGMTServer}
                        "\[$takeKeyword]" {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active$" -computername $scsmMGMTServer}}
                        "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "[I][R][0-9]+")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem $message "ir" $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk -computername $scsmMGMTServer}catch{}}}
                        {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                } 
        "sr" {
                    $workItem = get-scsmobject -class $srClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer| ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                        $assignedToSMTP.TargetAddress {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                        default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                    }
                    switch -Regex ($commentToAdd)
                    {
                        "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Completed$" -computername $scsmMGMTServer}
                        "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Cancelled$" -computername $scsmMGMTServer}
                        "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Closed$" -computername $scsmMGMTServer}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                } 
        "pr" {
                    $workItem = get-scsmobject -class $prClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {Add-ProblemComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                        default {Add-ProblemComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $null}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                         "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Resolved$" -computername $scsmMGMTServer; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                         "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Closed$" -computername $scsmMGMTServer}
                         "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                         "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Active$" -computername $scsmMGMTServer}}
                        {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                }
        "cr" {
                    $workItem = get-scsmobject -class $crClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try{$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {Add-ChangeRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                        default {Add-ChangeRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -isPrivate $false}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                        "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" -computername $scsmMGMTServer}
                        "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" -computername $scsmMGMTServer}
                        "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                }
        
        #### activities ####
        "ra" {
                    $workItem = get-scsmobject -class $raClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    $reviewers = Get-SCSMRelatedObject -SMObject $workItem -Relationship $raHasReviewerRelClass -computername $scsmMGMTServer
                    foreach ($reviewer in $reviewers)
                    {
                        $reviewingUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $reviewer -Relationship $raReviewerIsUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer
                        $reviewingUserSMTP = Get-SCSMRelatedObject -SMObject $reviewingUser -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress
                        
                        #approved
                        if (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$approvedKeyword]"))
                        {
                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} -computername $scsmMGMTServer
                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk -computername $scsmMGMTServer
                        }
                        #rejected
                        elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                        {
                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} -computername $scsmMGMTServer
                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk -computername $scsmMGMTServer
                        }
                        #no keyword, add a comment to parent work item
                        elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$approvedKeyword]") -or ($commentToAdd -notmatch "\[$rejectedKeyword]")))
                        {
                            $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                            switch ($parentWorkItem.Classname)
                            {
                                "System.WorkItem.ChangeRequest" {Add-ChangeRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.ServiceRequest" {Add-ServiceRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.Incident" {Add-IncidentComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            }
                            
                        }
                    }
                }
        "ma" {
                    $workItem = get-scsmobject -class $maClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$activityImplementer = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if ($activityImplementer){$activityImplementerSMTP = Get-SCSMRelatedObject -SMObject $activityImplementer -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    
                    #completed
                    if (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$completedKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #skipped
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$skipKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #not from the Activity Implementer, add to the MA Notes
                    elseif (($activityImplementerSMTP.TargetAddress -ne $message.From))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #no keywords, add to the Parent Work Item
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$completedKeyword]") -or ($commentToAdd -notmatch "\[$skipKeyword]")))
                    {
                        $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                        switch ($parentWorkItem.Classname)
                        {
                            "System.WorkItem.ChangeRequest" {Add-ChangeRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            "System.WorkItem.ServiceRequest" {Add-ServiceRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            "System.WorkItem.Incident" {Add-IncidentComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                        }
                            
                    }
                }
    } 
}

function Attach-EmailToWorkItem ($message, $workItemID)
{
    $messageMime = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $MemoryStream = New-Object System.IO.MemoryStream($messageMime.MimeContent.Content,0,$messageMime.MimeContent.Content.Length)

    #Create the attachment object itself and set its properties for SCSM
    $emailAttachment = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
    $emailAttachment.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
    $emailAttachment.Item($fileAttachmentClass, "DisplayName").Value = "message.eml"
    $emailAttachment.Item($fileAttachmentClass, "Description").Value = "ExchangeConversationID:$($message.ConversationID);"
    $emailAttachment.Item($fileAttachmentClass, "Extension").Value =   "eml"
    $emailAttachment.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
    $emailAttachment.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
    $emailAttachment.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
    
    #Add the attachment to the work item and commit the changes
    $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemID'" -computername $scsmMGMTServer
    $WorkItemProjection.__base.Add($emailAttachment, $fileAttachmentRelClass.Target)
    $WorkItemProjection.__base.Commit()
            
    #create the Attached By relationship if possible
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
    if ($userSMTPNotification) 
    { 
        $attachedByUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
        New-SCSMRelationshipObject -Source $emailAttachment -Relationship $fileAddedByUserRelClass -Target $attachedByUser -computername $scsmMGMTServer -Bulk
    }
}

#inspired and modified from Stefan Roth here - https://stefanroth.net/2015/03/28/scsm-passing-attachments-via-web-service-e-g-sma-web-service/
function Attach-FileToWorkItem ($message, $workItemId)
{
    foreach ($attachment in $message.Attachments)
    {
        if ($attachment.gettype().BaseType.Name -like "Mime*")
        {
            $signedAttachArray = $attachment.ContentObject.Stream.ToArray()
            $base64attachment = [System.Convert]::ToBase64String($signedAttachArray)
            $AttachmentContent = [convert]::FromBase64String($base64attachment)

            #Create a new MemoryStream object out of the attachment data
            $MemoryStream = New-Object System.IO.MemoryStream($signedAttachArray,0,$signedAttachArray.Length)

            if (([int]$MemoryStream.Length) -gt ($minFileSizeInKB+"kb"))
            {
                #Create the attachment object itself and set its properties for SCSM
                $NewFile = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
                $NewFile.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
                $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.FileName
                #$NewFile.Item($fileAttachmentClass, "Description").Value = $attachment.Description
                #$NewFile.Item($fileAttachmentClass, "Extension").Value =   $attachment.Extension
                $NewFile.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
                $NewFile.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
                $NewFile.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
    
                #Add the attachment to the work item and commit the changes
                $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemId'" -computername $scsmMGMTServer
                $WorkItemProjection.__base.Add($NewFile, $fileAttachmentRelClass.Target)
                $WorkItemProjection.__base.Commit()

                #create the Attached By relationship if possible
                $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
                if ($userSMTPNotification) 
                { 
                    $attachedByUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
                    New-SCSMRelationshipObject -Source $NewFile -Relationship $fileAddedByUserRelClass -Target $attachedByUser -computername $scsmMGMTServer -Bulk
                }
            }
        }
        else
        {
            $attachment.Load()
            $base64attachment = [System.Convert]::ToBase64String($attachment.Content)

            #Convert the Base64String back to bytes
            $AttachmentContent = [convert]::FromBase64String($base64attachment)

            #Create a new MemoryStream object out of the attachment data
            $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)

            if (([int]$MemoryStream.Length) -gt ($minFileSizeInKB+"kb"))
            {
                #Create the attachment object itself and set its properties for SCSM
                $NewFile = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
                $NewFile.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
                $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.Name
                #$NewFile.Item($fileAttachmentClass, "Description").Value = $attachment.Description
                #$NewFile.Item($fileAttachmentClass, "Extension").Value =   $attachment.Extension
                $NewFile.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
                $NewFile.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
                $NewFile.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
    
                #Add the attachment to the work item and commit the changes
                $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemId'" -computername $scsmMGMTServer
                $WorkItemProjection.__base.Add($NewFile, $fileAttachmentRelClass.Target)
                $WorkItemProjection.__base.Commit()

                #create the Attached By relationship if possible
                $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
                if ($userSMTPNotification) 
                { 
                    $attachedByUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
                    New-SCSMRelationshipObject -Source $NewFile -Relationship $fileAddedByUserRelClass -Target $attachedByUser -computername $scsmMGMTServer -Bulk
                }
            }
        }
    }
}

function Get-WorkItem ($workItemID, $workItemClass)
{
    #removes [] surrounding a Work Item ID if neccesary
    if ($workitemID.StartsWith("[") -and $workitemID.EndsWith("]"))
    {
        $workitemID = $workitemID.TrimStart("[").TrimEnd("]")
    }

    #get the work item
    $wi = get-scsmobject -Class $workItemClass -Filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
    return $wi
}

#courtesy of Leigh Kilday. Modified.
function Get-SCSMWorkItemParent
{
    [CmdLetBinding()]
    PARAM (
        [Parameter(ParameterSetName = 'GUID', Mandatory=$True)]
        [Alias('ID')]
        $WorkItemGUID
    )
    PROCESS
    {
        TRY
        {
            If ($PSBoundParameters['WorkItemGUID'])
            {
                Write-Verbose -Message "[PROCESS] Retrieving WI with GUID"
                $ActivityObject = Get-SCSMObject -Id $WorkItemGUID -computername $scsmMGMTServer
            }
        
            #Retrieve Parent
            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name)"
            Write-Verbose -Message "[PROCESS] Retrieving WI Parent"
            $ParentRelatedObject = Get-SCSMRelationshipObject -ByTarget $ActivityObject -computername $scsmMGMTServer | ?{$_.RelationshipID -eq $wiContainsActivityRelClass.id.Guid}
            $ParentObject = $ParentRelatedObject.SourceObject

            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name) - Parent: $($ParentObject.Name)"

            If ($ParentObject.ClassName -eq 'System.WorkItem.ServiceRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ChangeRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ReleaseRecord' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Incident' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Problem')
            {
                Write-Verbose -Message "[PROCESS] This is the top level parent"
                
                #return parent object Work Item
                Return $ParentObject
            }
            Else
            {
                Write-Verbose -Message "[PROCESS] Not the top level parent. Running against this object"
                Get-SCSMWorkItemParent -WorkItemGUID $ParentObject.Id.GUID -computername $scsmMGMTServer
            }
        }
        CATCH
        {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
function Create-UserInCMDB ($userEmail)
{
    #The ID for external users appears to be a GUID, but it can't be identified by get-scsmobject
    #The ID for internal domain users takes the form of domain_username_SMTP
    #It's unclear how this ID should be generated. Opted to take the form of an internal domain for the ID
    #By using the internal domain style (_SMTP) this means New/Update Work Item tasks will understand how to find these new external users going forward
    $username = $userEmail.Split("@")[0]
    $domainAndTLD = $userEmail.Split("@")[1]
    $domain = $domainAndTLD.Split(".")[0]
    $newID = $domain + "_" + $username + "_SMTP"

    #create the new user
    $newUser = New-SCSMObject -Class $domainUserClass -PropertyHashtable @{"domain" = "$domainAndTLD"; "username" = "$username"; "displayname" = "$userEmail"} -PassThru

    #create the user notification projection
    $userNoticeProjection = @{__CLASS = "$($domainUserClass.Name)";
                                __SEED = $newUser;
                                Notification = @{__CLASS = "$($notificationClass)";
                                                    __OBJECT = @{"ID" = $newID; "TargetAddress" = "$userEmail"; "DisplayName" = "E-mail address"; "ChannelName" = "SMTP"}
                                                }
                                }

    #create the user's email notification channel
    New-SCSMObjectProjection -Type "$($userHasPrefProjection.Name)" -Projection $userNoticeProjection

    return $newUser
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
function Add-IncidentComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the WI Object it passed to the function
    If ($WIObject.Id -ne $NULL) {

        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystComments"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "UserComments"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.ClassName)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Incident Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ServiceRequestComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystCommentLog"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "EndUserCommentLog"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Service Request Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ProblemComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "Comment"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "EndUserCommentLog"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Problem Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ChangeRequestComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystComments"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "UserComments"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        #NOTE: This Projection is 100% based on Cireson's CR projection as this is the ONLY projection
        #that features AssignedTo, AffectedUser, CreatedBy, and EndUser/Analyst Action Log comments
        #If you aren't a customer of Cireson, you'll need to create your own type projection
        #to use here.
        New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Change Request Object!"
    }
}

#retrieve a user from SCSM through the Cireson Web Portal API
function Get-CiresonPortalUser ($username, $domain)
{
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $isAuthUserAPIurl = "api/V3/User/IsUserAuthorized?userName=$username&domain=$domain"
        $returnedUser = Invoke-WebRequest -Uri ($ciresonPortalServer+$isAuthUserAPIurl) -Method post -UseDefaultCredentials -SessionVariable userWebRequestSessionVar
        $ciresonPortalUserObject = $returnedUser.Content | ConvertFrom-Json
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable userWebRequestSessionVar
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -Method post -Body $loginForm.Fields -WebSession $userWebRequestSessionVar 
        $isAuthUserAPIurl = "api/V3/User/IsUserAuthorized?userName=$username&domain=$domain"
        $returnedUser = Invoke-WebRequest -Uri ($ciresonPortalServer+$isAuthUserAPIurl) -Method post -WebSession $userWebRequestSessionVar
        $ciresonPortalUserObject = $returnedUser.Content | ConvertFrom-Json
    }

    return $ciresonPortalUserObject
}

#retrieve a group from SCSM through the Cireson Web Portal API
function Get-CiresonPortalGroup ($groupEmail)
{
    $groupName = Get-ADGroup -Filter "Mail -eq $groupEmail"

    if($ciresonPortalWindowsAuth)
    {
        #wanted to use a get groups style request, but "api/V3/User/GetConsoleGroups" feels costly instead of a search
        $cwpGroupResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+"api/V3/User/GetUserList?userFilter=$($groupName.Name)&filterByAnalyst=false&groupsOnly=true&maxNumberOfResults=25") -UseDefaultCredentials
        $ciresonPortalGroup = ($cwpGroupResponse.content | ConvertFrom-Json) | select-object @{Name='AccessGroupId'; Expression={$_.Id}}, name | ?{$_.name -eq $($groupName.Name)} 
        return $ciresonPortalGroup
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable groupWebRequestSessionVar
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -Method post -Body $loginForm.Fields -WebSession $groupWebRequestSessionVar 
        $cwpGroupResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+"api/V3/User/GetUserList?userFilter=$($groupName.Name)&filterByAnalyst=false&groupsOnly=true&maxNumberOfResults=25") -WebSession $groupWebRequestSessionVar
        $ciresonPortalGroup = ($cwpGroupResponse.content | ConvertFrom-Json) | select-object @{Name='AccessGroupId'; Expression={$_.Id}}, name | ?{$_.name -eq $($groupName.Name)} 
        return $ciresonPortalGroup
    }
}

#retrieve all the announcements on the portal
function Get-CiresonPortalAnnouncements ($languageCode)
{
    if($ciresonPortalWindowsAuth)
    {
        $allAnnouncementsURL = "api/V3/Announcement/GetAllAnnouncements?languageCode=$($languageCode)"
        $allCiresonPortalAnnouncements = Invoke-WebRequest -uri ($ciresonPortalServer+$allAnnouncementsURL) -UseDefaultCredentials | ConvertFrom-Json
        return $allCiresonPortalAnnouncements
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable announcementWebRequestSessionVar
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -Method post -Body $loginForm.Fields -WebSession $announcementWebRequestSessionVar
        $allAnnouncementsURL = "api/V3/Announcement/GetAllAnnouncements?languageCode=$($languageCode)"
        $allCiresonPortalAnnouncements = Invoke-WebRequest -uri ($ciresonPortalServer+$allAnnouncementsURL) -WebSession $announcementWebRequestSessionVar | ConvertFrom-Json
        return $allCiresonPortalAnnouncements
    }
}

#search for available Request Offerings based on content from a New Work Item and notify the Affected user via the Cireson Portal API
function Search-AvailableCiresonPortalOfferings ($workItem, $ciresonPortalUser)
{
    $searchQuery = $workItem.Title.Trim() + " " + $workItem.Description.Trim()

    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $serviceCatalogAPIurl = "api/V3/ServiceCatalog/GetServiceCatalog?userId=$($ciresonPortalUser.id)&isScoped=$($ciresonPortalUser.Security.IsServiceCatalogScoped)"
        $serviceCatalogResults = Invoke-WebRequest -Uri ($ciresonPortalServer+$serviceCatalogAPIurl) -Method get -UseDefaultCredentials -SessionVariable ecWebSession
        $serviceCatalogResults = $serviceCatalogResults.Content | ConvertFrom-Json
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable ecWebSession
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -Method post -Body $loginForm.Fields -WebSession $ecWebSession 
        $serviceCatalogAPIurl = "api/V3/ServiceCatalog/GetServiceCatalog?userId=$($ciresonPortalUser.id)&isScoped=$($ciresonPortalUser.Security.IsServiceCatalogScoped)"
        $serviceCatalogResults = Invoke-WebRequest -Uri ($ciresonPortalServer+$serviceCatalogAPIurl) -Method get -WebSession $ecWebSession
        $serviceCatalogResults = $serviceCatalogResults.Content | ConvertFrom-Json | Select-Object RequestOfferingTitle, RequestOfferingDescription, Service, RequestOfferingId, ServiceOfferingId
    }

    #### If the user has access to some Request Offerings, find which RO Titles/Description contain words from their original message ####
    if ($serviceCatalogResults)
    {
        #prepare the results by generating a URL array for the email
        $matchingRequestURLs = @()
        foreach ($serviceCatalogResult in $serviceCatalogResults)
        {
            $wordsMatched = ($searchQuery.Split() | ?{($serviceCatalogResult.title -match "\b$_\b") -or ($serviceCatalogResult.description -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToRO)
            {
                $ciresonPortalRequestURL = "`"" + $ciresonPortalServer + "SC/ServiceCatalog/RequestOffering/" + $serviceCatalogResult.RequestOfferingId + "," + $serviceCatalogResult.ServiceOfferingId + "`""
                $matchingRequestURLs += "<a href=$ciresonPortalRequestURL/>$($serviceCatalogResult.RequestOfferingTitle)</a><br />"
            }
        }

        return $matchingRequestURLs
    }
}

#search the Cireson KB based on content from a New Work Item and notify the Affected User
function Search-CiresonKnowledgeBase ($workItem, $ciresonPortalUser)
{
    $searchQuery = $workItem.Title.Trim() + " " + $workItem.Description.Trim()

    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$($ciresonPortalUser.Id)&searchValue=$searchQuery&isManager=$([bool]$ciresonPortalUser.KnowledgeManager)&userLanguageCode=$($ciresonPortalUser.LanguageCode)") -UseDefaultCredentials
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable ecPortalSession
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -WebSession $ecPortalSession -Method post -Body $loginForm.Fields
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$($ciresonPortalUser.Id)&searchValue=$searchQuery&isManager=$([bool]$ciresonPortalUser.KnowledgeManager)&userLanguageCode=$($ciresonPortalUser.LanguageCode)") -WebSession $ecPortalSession
    }

    $kbResults = $kbResults.Content | ConvertFrom-Json
    $kbResults =  $kbResults | ?{$_.endusercontent -ne ""} | select-object articleid, title
    
    if ($kbResults)
    {
        $matchingKBURLs = @()
        foreach ($kbResult in $kbResults)
        {
            $matchingKBURLs += "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
        }

        return $matchingKBURLs
    }
}

#send an email from the SCSM Workflow Account
function Send-EmailFromWorkflowAccount ($subject, $body, $bodyType, $toRecipients)
{
    $emailToSendOut = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchangeService
    $emailToSendOut.Subject = $subject
    $emailToSendOut.Body = $body
    $emailToSendOut.ToRecipients.Add($toRecipients)
    $emailToSendOut.Body.BodyType = $bodyType
    $emailToSendOut.Send()
}

function Schedule-WorkItem ($calAppt, $wiType, $workItem)
{
    if ($calAppt.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        #set the Scheduled Start/End dates on the Work Item
        $scheduledHashTable =  @{"ScheduledStartDate" = $calAppt.StartTime.ToUniversalTime(); "ScheduledEndDate" = $calAppt.EndTime.ToUniversalTime()}    
        
        switch ($wiType)
        {
            "ir" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "sr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "pr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "cr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "rr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}

            #activities
            "ma" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "pa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "sa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "da" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        }
    }

    #the meeting request is a cancelation, null the values out of the Schedule Start/End Time
    elseif ($calAppt.ItemClass -eq "IPM.Schedule.Meeting.Canceled")
    {
        #set the Scheduled Start/End dates on the Work Item
        $scheduledHashTable =  @{"ScheduledStartDate" = $null; "ScheduledEndDate" = $null}    
        
        switch ($wiType)
        {
            "ir" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "sr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "pr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "cr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "rr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}

            #activities
            "ma" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "pa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "sa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
            "da" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        }
    }
}

function Verify-WorkItem ($message)
{
    #If emails are being attached to New Work Items, filter on the File Attachment Description that equals the Exchange Conversation ID as defined in the Attach-EmailToWorkItem function
    if ($attachEmailToWorkItem -eq $true)
    {
        $emailAttachmentSearchObject = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" -ComputerName $scsmMGMTServer | select-object -first 1 
        $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $emailAttachmentSearchObject -ComputerName $scsmMGMTServer).sourceobject.id -ComputerName $scsmMGMTServer
        if ($emailAttachmentSearchObject -and $relatedWorkItemFromAttachmentSearch)
        {
            switch ($relatedWorkItemFromAttachmentSearch.ClassName)
            {
                "System.WorkItem.Incident" {Update-WorkItem -message $message -wiType "ir" -workItemID $relatedWorkItemFromAttachmentSearch.id}
                "System.WorkItem.ServiceRequest" {Update-WorkItem -message $message -wiType "sr" -workItemID $relatedWorkItemFromAttachmentSearch.id}
            }
        }
        else
        {
            #no match was found, Create a New Work Item
            New-WorkItem $message $defaultNewWorkItem
        }
    }
    else
    {
        #will never engage as Verify-WorkItem currently only works when attaching emails to work items 
    }
}

function Read-MIMEMessage ($message)
{
    #Get the Mime Content of the message via MimeKit
    $messageWithMimeContent = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $mimeMessageMemoryStream = New-Object System.IO.MemoryStream($messageWithMimeContent.MimeContent.Content,0,$messageWithMimeContent.MimeContent.Content.Length)
    $parsedMimeMessage = New-Object MimeKit.MimeParser($mimeMessageMemoryStream)

    return $parsedMimeMessage
}

#retrieve sender's ability to post announcement based on previously defined email addresses or an AD group
function Get-SCSMAuthorizedAnnouncer ($sender)
{
    switch ($approvedMemberTypeForSCSMAnnouncer)
    {
        "users" {if ($approvedUsersForSCSMAnnouncements -match $sender)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }        
                }
        "group" {$group = Get-ADGroup $approvedADGroupForSCSMAnnouncements
                    $adUser = Get-ADUser -Filter "EmailAddress -eq '$sender'"
                    if ($adUser)
                    {
                        if ((Get-ADUser $adUser -Properties memberof).memberof -eq $group.distinguishedname)
                        {
                            return $true
                        }
                        else
                        {
                            return $false
                        }
                    }
                }
    }
}

function Set-CoreSCSMAnnouncement ($message, $workItem)
{
    #if the message is an email, we need to add the end time property to the object
    #otherwise, it's a calendar appointment/meeting which already has these properties
    if ($message.ItemClass -eq "IPM.Note")
    {
        $message | Add-Member -type NoteProperty -name StartTime -Value $message.DateTimeReceived
        $message | Add-Member -type NoteProperty -name EndTime -Value $null
    }

    #Parse the message body for a Priority #keyword to correlate to a SCSM Priority
    #Rearrange the keyword from the title body before posting the announcement
    #If the end time is null, generate it
    $announcementTitle = $message.Subject -replace "\[$($workItem.Name)\]", ""
    $announcementTitle = $announcementTitle + " " + "[$($workItem.Name)]"
    if ($message.body -match [Regex]::Escape("#$lowAnnouncemnentPriorityKeyword"))
    {
        #low priority
        $scsmPriorityName = "Low"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "\#$lowAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
    }
    elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
    {
        #high priority
        $scsmPriorityName = "Critical"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
    }
    else
    {
        #normal priority
        $scsmPriorityName = "Medium"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
    }

    $announcementClass = get-scsmclass "System.Announcement.Item$" -ComputerName $scsmMGMTServer
    $announcementPropertyHashtable = @{"Title" = $announcementTitle; "Body" = $announcementBody; "ExpirationDate" = $message.EndTime.ToUniversalTime(); "Priority" = $scsmPriorityName}

    #get any current announcement to update, otherwise create
    $currentSCSMAnnouncements = Get-SCSMObject -Class $announcementClass -Filter "Title -like '*$($workitem.Name)*'" -ComputerName $scsmMGMTServer

    if ($currentSCSMAnnouncements)
    {
        foreach ($currentSCSMAnnouncement in $currentSCSMAnnouncements)
        {
            Set-SCSMObject -SMObject $currentSCSMAnnouncement -PropertyHashtable $announcementPropertyHashtable -ComputerName $scsmMGMTServer
        }
    }
    else
    {
        #create the announcement in SCSM
        New-SCSMObject -Class $announcementClass -PropertyHashtable $announcementPropertyHashtable -ComputerName $scsmMGMTServer
    }
}

function Set-CiresonPortalAnnouncement ($message, $workItem)
{
    $updateAnnouncementURL = "api/v3/Announcement/UpdateAnnouncement"

    #if the message is an email, we need to add the start time and end time property to the object
    #otherwise, it's a calendar appointment/meeting which already has these properties
    if ($message.ItemClass -eq "IPM.Note")
    {
        $message | Add-Member -type NoteProperty -name StartTime -Value $message.DateTimeReceived
        $message | Add-Member -type NoteProperty -name EndTime -Value $null
    }

    #Parse the message body for a Priority #keyword to correlate to a Cireson Priority enum
    #Remove the keyword from the title body before posting the announcement
    #If the end time is null, generate it
    $announcementTitle = $message.Subject -replace "\[$($workItem.Name)\]", ""
    $announcementTitle = $announcementTitle + " " + "[$($workItem.Name)]"
    if ($message.body -match [Regex]::Escape("#$lowAnnouncemnentPriorityKeyword"))
    {
        #low priority
        $ciresonPortalPriorityEnum = "F860661B-D9D6-41CB-A501-467B4DD81A7B"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "\#$lowAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
    }
    elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
    {
        #high priority
        $ciresonPortalPriorityEnum = "F10A51C2-C569-4E64-8237-2B117D63DDB8"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
    }
    else
    {
        #normal priority
        $ciresonPortalPriorityEnum = "64096F7F-F8E0-491C-A7FE-94FEDDED4715"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
    }

    #Extract the groups that the message was sent to
    #rename the GroupID property to "AccessGroupID" so as to compare the difference later
    $groupEmails = @()
    $groupEmails += $message.To | ?{$_.MailboxType -ne "Mailbox"}
    $groupEmails += $message.Cc | ?{$_.MailboxType -ne "Mailbox"}
    $portalGroups = @()
    foreach ($groupEmail in $groupEmails)
    {
        $portalGroups += Get-CiresonPortalGroup -groupEmail $groupEmail.Name
    }

    #Get the user that is posting the announcement (from) from the SCSM/Cireson Portal to determine their language code to post the announcement
    $announcerSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer | sort-object lastmodified -Descending | select -first 1
    $announcerSCSMObject = Get-SCSMObject -id (Get-SCSMRelationshipObject -ByTarget $announcerSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    $ciresonPortalAnnouncer = Get-CiresonPortalUser -username $announcerSCSMObject.username -domain $announcerSCSMObject.domain

    #Get any announcements that already exist for the Work Item
    $allPortalAnnouncements = Get-CiresonPortalAnnouncements -languageCode $ciresonPortalAnnouncer.LanguageCode
    $allPortalAnnouncements = $allPortalAnnouncements | ?{$_.title -match "\[" + $workitem.name + "\]"}

    #determine authentication to use (windows/forms)
    if ($ciresonPortalWindowsAuth -eq $false)
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable newAnnouncementWebRequestSessionVar
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "Login/Login?ReturnUrl=%2f") -Method post -Body $loginForm.Fields -WebSession $newAnnouncementWebRequestSessionVar
    }

    if ($allPortalAnnouncements)
    {
        #### there are announcements to create/update ####

        #### combine the announcement objects and group objects together and group by GroupAccessID, then find object groups that don't have an announcement id ####
        #Announcement array has an AccessGroupID that does not match a group from the message. create announcement for that group
        $groupsToCreateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | ?{$_.Count -eq 1} | Select-Object -Expand Group | ?{$_.Id -eq $null}

        #Announcement array has an AccessGroupID that contains a current group from the message.
        $groupsToUpdateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | Select-Object -Expand Group | ?{$_.Id -ne $null}

        # create announcement for new group
        foreach ($groupsToCreateAnnouncement in $groupsToCreateAnnouncements)
        {
            #create the Portal Announcement Hashtable, convert to JSON object, then POST to create
            $announcement = @{"Id" = [guid]::NewGuid();
                                "Title" = $announcementTitle;
                                "Body" = $announcementBody;
                                "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                "AccessGroupId" = @{"Id"=$($groupsToCreateAnnouncement.AccessGroupId)};
                                "StartDate" = $message.StartTime.ToUniversalTime();
                                "EndDate" = $message.EndTime.ToUniversalTime();
                                "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
            $announcement = $announcement | ConvertTo-Json

            #post the announcement
            if ($ciresonPortalWindowsAuth)
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -WebSession $newAnnouncementWebRequestSessionVar
            }
        }

        # update current announcement's title and body
        foreach ($groupsToUpdateAnnouncement in $groupsToUpdateAnnouncements)
        {
            #create the Portal Announcement Hashtable using current Announcement ID, convert to JSON object, then POST to update
            $announcement = @{"Id" = $groupsToUpdateAnnouncement.Id;
                                "Title" = $announcementTitle;
                                "Body" = $announcementBody;
                                "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                "AccessGroupId" = @{"Id"=$($groupsToUpdateAnnouncement.AccessGroupId)};
                                "StartDate" = $message.StartTime.ToUniversalTime();
                                "EndDate" = $message.EndTime.ToUniversalTime();
                                "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
            $announcement = $announcement | ConvertTo-Json

            #post the announcement
            if ($ciresonPortalWindowsAuth)
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -WebSession $newAnnouncementWebRequestSessionVar
            }
        }
    }
    else
    {
        #### there are announcements to create ####

        #Cireson Portal Announcements can only target a single group. Create an announcement for each group
        foreach ($portalGroup in $portalGroups)
        {
            #create the Portal Announcement Hashtable, convert to JSON object, then POST
            $announcement = @{"Id" = [guid]::NewGuid();
                                "Title" = $announcementTitle;
                                "Body" = $announcementBody;
                                "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                "AccessGroupId" = @{"Id"=$($portalGroup.AccessGroupId)};
                                "StartDate" = $message.StartTime.ToUniversalTime();
                                "EndDate" = $message.EndTime.ToUniversalTime();
                                "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
            $announcement = $announcement | ConvertTo-Json

            #post the announcement
            if ($ciresonPortalWindowsAuth)
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-WebRequest -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -WebSession $newAnnouncementWebRequestSessionVar
            }
        }
    }
}
#endregion

#region #### SCOM Request Functions ####
function Get-SCOMAuthorizedRequester ($sender)
{
    switch ($approvedMemberTypeForSCOM)
    {
        "users" {if ($approvedUsersForSCOM -match $sender)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }        
                }
        "group" {$group = Get-ADGroup $approvedADGroupForSCOM
                    $adUser = Get-ADUser -Filter "EmailAddress -eq '$sender'"
                    if ($adUser)
                    {
                        if ((Get-ADUser $adUser -Properties memberof).memberof -eq $group.distinguishedname)
                        {
                            return $true
                        }
                        else
                        {
                            return $false
                        }
                    }
                }
    }
}

function Get-SCOMDistributedAppHealth ($message)
{
    #determine if the sender is authorized to make SCOM Health requests
    $isAuthorized = Get-SCOMAuthorizedRequester $message.From.Address

    if (($isAuthorized -eq $true))
    {
        #find the distributed application to search for based on the [Distributed App Name] from the email body
        #"\[(.*?)\]" - will match something [Service Manager] or [Operations Manager Management Group]
        if ($message.body -match "\[(.*?)\]"){$appName = $Matches[0].Replace("[", "").Replace("]", "")}
        else {<#body not [formed] correctly#>}

        #get Distributed Applications that meet search criteria
        $distributedApps = invoke-command -scriptblock {(Get-SCOMClass | Where-Object {$_.displayname -like "*$appName*"} | Get-SCOMMonitoringObject) | select-object Displayname, healthstate} -ComputerName $scomMGMTServer
        $healthySCOMApps = @()
        $unhealthySCOMApps = @()
        $notMonitoredSCOMApps = @()
        $unhealthySCOMAppsAlerts = @()
        $emailBody = @()

        #create, define, and load custom PS Object from SCOM DA Objects
        foreach ($distributedApp in $distributedApps)
        {
            #Healthy app - Green Agent state
            if ($distributedApp.HealthState -eq "Success")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Healthy"
                $healthySCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            }
            #Unhealthy App - Red Agent state
            elseif ($result.HealthState -eq "Error")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Critical"
                $unhealthySCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            }
            #Gray Agent state
            elseif ($result.HealthState -eq "Uninitialized")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Not Monitored"
                $notMonitoredSCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            } 
        }

        #if there are unhealthy apps/red agent states, get their Active alerts in SCOM
        if ($unhealthySCOMApps)
        {
            foreach ($unhealthySCOMApp in $unhealthySCOMApps)
            {
                $unhealthySCOMAppsAlerts += invoke-command -scriptblock {Get-SCOMClass | Where-Object {$_.displayname -like “$($unhealthySCOMApp.displayname)”} | Get-SCOMClassInstance | %{$_.GetRelatedMonitoringObjects()} | Get-SCOMAlert -ResolutionState 0} -computername $scomMGMTServer
            }
        }
        
        $emailBody = $emailBody + "<br /><br />" + "NOTE: Responding to this message will trigger the creation of a New Work Item in Service Manager!"
        Send-EmailFromWorkflowAccount -subject "SCOM Health Status" -body $emailBody -bodyType "HTML" -toRecipients $message.From
    }
    else
    {
        return $false
    }
}
#endregion

#determine/enforce merge logic in the event this was omitted in configuration
if (($mergeReplies -eq $true) -or ($processCalendarAppointment -eq $true))
{
    $attachEmailToWorkItem = $true
}

#load the MimeKit assembly and decrypting context/certificate
if (($processDigitallySignedMessages -eq $true) -or ($processEncryptedMessages -eq $true))
{
    try
    {
        [void][System.Reflection.Assembly]::LoadFile($mimeKitDLLPath)
        if ($certStore -eq "user")
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("CurrentUser")
        }
        else
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("LocalMachine")
        }
    }
    catch
    {
        #decrypting certificate or mimekit couldn't be loaded. Don't process signed/encrypted emails
        $processDigitallySignedMessages = $false
        $processEncryptedMessages = $false
    }
}

#define Exchange assembly and connect to EWS
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
switch ($exchangeAuthenticationType)
{
    "impersonation" {$exchangeService.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)}
    "windows" {$exchangeService.UseDefaultCredentials = $true}
}
$exchangeService.AutodiscoverUrl($workflowEmailAddress)

#define search parameters and search on the defined classes
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
$inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
$mimeContentSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
$dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
$now = get-date
$searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$now

#build the Where-Object scriptblock based on defined configuration
#by default the connector will ALWAYS process regular emails as seen in the $emailFilterString variable
$emailFilterString = '($_.ItemClass -eq "IPM.Note")'
$calendarFilterString = '($_.ItemClass -eq "IPM.Schedule.Meeting.Request") -or ($_.ItemClass -eq "IPM.Schedule.Meeting.Canceled")'
$digitallySignedFilterString = '($_.ItemClass -eq "IPM.Note.SMIME.MultipartSigned")'
$encryptedFilterString = '($_.ItemClass -eq "IPM.Note.SMIME")'
$unreadFilterString = '($_.isRead -eq $false)'
$inboxFilterString = @()
if ($processCalendarAppointment -eq $true)
{
    $inboxFilterString += $calendarFilterString
}
if ($processDigitallySignedMessages -eq $true)
{
    $inboxFilterString += $digitallySignedFilterString
}
if ($processEncryptedMessages -eq $true)
{
    $inboxFilterString += $encryptedFilterString
}

#finalize the where-object string by ensuring to look for all Unread Items
$inboxFilterString = $inboxFilterString -join ' -or '
$inboxFilterString = "(" + $inboxFilterString + " -or " +  $emailFilterString + ")" + " -and " + $unreadFilterString
$inboxFilterString = [scriptblock]::Create("$inboxFilterString")

#filter the inbox
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView) | where-object $inboxFilterString | Sort-Object DateTimeReceived

#parse each message
foreach ($message in $inbox)
{
    #load the entire message
    $message.Load($propertySet)

    #Process an Email
    if ($message.ItemClass -eq "IPM.Note")
    {
        $email = New-Object System.Object 
        $email | Add-Member -type NoteProperty -name From -value $message.From.Address
        $email | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $email | Add-Member -type NoteProperty -name CC -value $message.CcRecipients
        $email | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $email | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $email | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $email | Add-Member -type NoteProperty -name ID -Value $message.ID
        $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
        $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####
            "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
    }

    #### Process a Digitally Signed message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME.MultipartSigned")
    {
        $response = Read-MIMEMessage $message

        #check to see if there are attachments besides the smime.p7s signature
        $signedAttachments = $response.Attachments
        $signedAttachments = $signedAttachments | ?{$_.filename -ne "smime.p7s"}
   
        $email = New-Object System.Object 
        $email | Add-Member -type NoteProperty -name From -value $response.From.address
        $email | Add-Member -type NoteProperty -name To -value $response.To.Address
        $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
        $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
        $email | Add-Member -type NoteProperty -name Attachments -value $signedAttachments
        $email | Add-Member -type NoteProperty -name Body -value $response.TextBody
        $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $email | Add-Member -type NoteProperty -name ID -Value $message.Id
        $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
        $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
        $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####
            "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
    }

    #### Process an Encrypted message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME")
    {
        $response = Read-MIMEMessage $message
        $decryptedBody = $response.Body.Decrypt($certStore)

        #Messaged is encrypted
        if (($response.Body -ne $null) -and ($response.Body.SecureMimeType -eq "EnvelopedData") -and ($decryptedBody.TextBody))
        {         
            #check to see if there are attachments
            $decryptedAttachments = $decryptedBody | ?{$_.isattachment -eq $true}

            $email = New-Object System.Object 
            $email | Add-Member -type NoteProperty -name From -value $response.From.Address
            $email | Add-Member -type NoteProperty -name To -value $response.To.Address
            $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
            $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
            $email | Add-Member -type NoteProperty -name Attachments -value $decryptedAttachments
            $email | Add-Member -type NoteProperty -name Body -value $decryptedBody.TextBody
            $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
            $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
            $email | Add-Member -type NoteProperty -name ID -Value $message.Id
            $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
            $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
            $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

            switch -Regex ($email.subject) 
            { 
                #### primary work item types ####
                "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
                #### activities ####
                "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
                "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

                #### default action, create work item ####
                default {new-workitem $email $defaultNewWorkItem} 
            }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
            $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
        }
        #Message is encrypted and signed
        else
        {
            $email = New-Object System.Object 
            $email | Add-Member -type NoteProperty -name From -value $response.From.Address
            $email | Add-Member -type NoteProperty -name To -value $response.To.Address
            $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
            $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
            $email | Add-Member -type NoteProperty -name Attachments -value $decryptedAttachments
            $email | Add-Member -type NoteProperty -name Body -value "This message is digitally encrypted and signed. Please see the related/attached item."
            $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
            $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
            $email | Add-Member -type NoteProperty -name ID -Value $message.Id
            $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
            $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
            $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

            switch -Regex ($email.subject) 
            { 
                #### primary work item types ####
                "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
                #### activities ####
                "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
                "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

                #### default action, create work item ####
                default {new-workitem $email $defaultNewWorkItem} 
            }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
            $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
        }
    }

    #Process a Calendar Meeting
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        $appointment = New-Object System.Object 
        $appointment | Add-Member -type NoteProperty -name StartTime -value $message.Start
        $appointment | Add-Member -type NoteProperty -name EndTime -value $message.End
        $appointment | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $appointment | Add-Member -type NoteProperty -name From -value $message.From.Address
        $appointment | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $appointment | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $appointment | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $appointment | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $appointment | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $appointment | Add-Member -type NoteProperty -name ID -Value $message.ID
        $appointment | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $appointment | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
        $appointment | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ir" -workItemID $result.name}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sr" -workItemID $result.name}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pr" -workItemID $result.name}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "cr" -workItemID $result.name}}
            "\[[R][R][0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "rr" -workItemID $result.name}}

            #### activities ####
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ma" -workItemID $result.name}}
            "\[[P][A][0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pa" -workItemID $result.name}}
            "\[[S][A][0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sa" -workItemID $result.name}}
            "\[[D][A][0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "da" -workItemID $result.name}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)} 
        }
    }

    #Process a Calendar Meeting Cancellation
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Canceled")
    {
        $appointment = New-Object System.Object 
        $appointment | Add-Member -type NoteProperty -name StartTime -value $message.Start
        $appointment | Add-Member -type NoteProperty -name EndTime -value $message.End
        $appointment | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $appointment | Add-Member -type NoteProperty -name From -value $message.From.Address
        $appointment | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $appointment | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $appointment | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $appointment | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $appointment | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $appointment | Add-Member -type NoteProperty -name ID -Value $message.ID
        $appointment | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $appointment | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
        $appointment | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass

        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; Update-WorkItem -message $appointment -wiType "ir" -workItemID $result.name}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; Update-WorkItem -message $appointment -wiType "sr" -workItemID $result.name}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; Update-WorkItem -message $appointment -wiType "pr" -workItemID $result.name}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; Update-WorkItem -message $appointment -wiType "cr" -workItemID $result.name}}
            "\[[R][R][0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; Update-WorkItem -message $appointment -wiType "rr" -workItemID $result.name}}

            #### activities ####
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; Update-WorkItem -message $appointment -wiType "ma" -workItemID $result.name}}
            "\[[P][A][0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; Update-WorkItem -message $appointment -wiType "pa" -workItemID $result.name}}
            "\[[S][A][0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; Update-WorkItem -message $appointment -wiType "sa" -workItemID $result.name}}
            "\[[D][A][0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; Update-WorkItem -message $appointment -wiType "da" -workItemID $result.name}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)} 
        }

        #Move to deleted items
        $message.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
    }
}
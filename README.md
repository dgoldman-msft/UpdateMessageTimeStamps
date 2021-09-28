# UpdateMessageTimeStamps

A helper script that allows you to change message time stamps for reproducing issues in a lab

## EXAMPLE

<span style="color:green">EXAMPLE 1</span>

> Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -UserImpersonation -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com

This will log on to the specified tenant using impersonation to a target mailbox and just report the messages in the mailbox

<span style="color:green">EXAMPLE 2</span>

> Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -UserImpersonation -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com -ChangingMessageStamps

This will log on to the specified tenant using impersonation to a target mailbox and just report the messages in the mailbox and change all the time stamps on a message (CreataionTime, SubmissionsTime, DeliveryTime and LastModificationTime)

<span style="color:green">EXAMPLE 3</span>

> Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com

This will log on to the specified tenant using not using impersonation and log in to the mailbox that the OAUTh token was generated for.

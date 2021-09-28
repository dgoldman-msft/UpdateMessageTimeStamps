# UpdateMessageTimeStamps

A helper script that allows you to change message time stamps for reproducing issues in a lab

## EXAMPLE

        Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -TenantName YourTenant@onmicrosoft.com

        This will log on to the tenant and just report the messages in the mailbox

        Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -TenantName YourTenant@onmicrosoft.com -ChangingMessageStamps

        This will log on to the tenant, connect to the mailbox and change all the time stamps on a message (CreataionTime, SubmissionsTime, DeliveryTime and LastModificationTime)

# Transport Agents
### This repository contains a Sample Transport Agent that implement various functions by leveraging Trasnport Agents in Microsoft Exchange 2019 CU12. 

## Features

- NDRAgent: Drops Non-Delivery Report (NDR) that contains in the body the word "DELETE"
- TaggingAgent: Redirect a message to somebody+text@domain to somebody@domain, implementing Plus Addressing
- HeaderAgent: Insert a custom header with a custom value in the header of every message that traverse the mail server
- RecipientDomainRewriteAgent: For any message sent to an address whose domain part is "contoso.com", it redirect the message to the same recipient on domain "tailspin.com"
- SenderDomainRewriteAgent: For any message received from an address whose domain part is "contoso.com", it changes the sending domain and make the message appear from "tailspin.com"

## Installation
1.	Copy the DLL to the server (i.e. F:\Transport Agents\\)
2.	Make sure the Exchange acconts have access to the folder
3.	Install-TransportAgent -Name <YourChosenName> -TransportAgentFactory "TransportAgent.<NameOfTheClass>" -AssemblyPath "F:\Transport Agents\TransportAgents.dll"
4.	Enable-TransportAgent <YourChosenName> 
5.	Exit from Exchange Management Shell
6.	Restart the MSExchangeTransport service

## Removal
1.	Disable-TransportAgent <YourChosenName>
2.	Uninstall-TransportAgent <YourChosenName>
3.	Exit from Exchange Management Shell
4.	Restart the MSExchangeTransport service

## Notes
More on Transport Agents can be found on https://learn.microsoft.com/en-us/exchange/mail-flow/transport-agents/transport-agents?view=exchserver-2019

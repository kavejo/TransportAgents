# Transport Agents
### This repository contains various Sample Transport Agents, fully functional and extensively tested, that achieve different business goals by leveraging various Trasnport Agents capabiltiies in Microsoft Exchange Server. 

#### This code is sample code. This sample is provided "as is" without warranty of any kind.

#### Microsoft and myslef further disclaims all implied warranties including without limitation any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the samples remains with you. In no event shall myself, Microsoft or its suppliers be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the samples, even if Microsoft has been advised of the possibility of such damages.

## Features

- [AutoResponderAgent](https://github.com/kavejo/TransportAgents/wiki/AutoResponderAgent): Sends an automatic response to emails directed to a mailbox or address that is being deprecated/removed
- [DomainReroutingAgent](https://github.com/kavejo/TransportAgents/wiki/DomainReroutingAgent): Rewrite the routing domain to be a differnt one (can be used to re-route traffic via a specific send connector matching the domain name space) when the sender is one of the configured ones
- [HeaderAgent](https://github.com/kavejo/TransportAgents/wiki/HeaderAgent): Insert a custom header with a custom value in the header of every message that traverse the mail server
- [HeaderReroutingAgent](https://github.com/kavejo/TransportAgents/wiki/HeaderReroutingAgent): Rewrite the routing domain to be a differnt one (can be used to re-route traffic via a specific send connector matching the domain name space) when a specific header is present in the message
- [InspectingAgent](https://github.com/kavejo/TransportAgents/wiki/InspectingAgent): Can be used to dump on event log information about the processed messages; this will write P1 (envelope), P2 (message) and header information
- [NDRAgent](https://github.com/kavejo/TransportAgents/wiki/NDRAgent): Drops Non-Delivery Report (NDR) that contains in the body the word "DELETE"
- [RecipientDomainRewriteAgent](https://github.com/kavejo/TransportAgents/wiki/RecipientDomainRewriteAgent): For any message sent to an address whose domain part is "contoso.com", it redirect the message to the same recipient on domain "tailspin.com"
- [SenderDomainRewriteAgent](https://github.com/kavejo/TransportAgents/wiki/SenderDomainRewriteAgent): For any message received from an address whose domain part is "contoso.com", it changes the sending domain and make the message appear from "tailspin.com"
- [TaggingAgent](https://github.com/kavejo/TransportAgents/wiki/TaggingAgent): Redirect a message to somebody+text@domain to somebody@domain, implementing Plus Addressing

## Wiki

For detailed information refer to the avaialble [Wiki](https://github.com/kavejo/TransportAgents/wiki)

## Note about Exchange Server versions

While this code specifically references Exhcange 2019 DLL's, it can be easily recompiled using Exchange 2013 or 2016 libraries, oiffering backward compatibility.

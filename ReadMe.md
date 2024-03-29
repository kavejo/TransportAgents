# Transport Agents
### This repository contains a Sample Transport Agent that implement various functions by leveraging Trasnport Agents in Microsoft Exchange 2019. 

## Features

- [AutoResponderAgent](https://github.com/kavejo/TransportAgents/wiki/AutoResponderAgent): Sends an automatic response to emails directed to a mailbox or address that is being deprecated/removed
- [DomainReroutingAgent](https://github.com/kavejo/TransportAgents/wiki/DomainReroutingAgent): Rewrite the routing domain to be a differnt one (can be used to re-route traffic via a specific send connector matching the domain name space) when the sender is one of the configured ones
- [HeaderAgent](https://github.com/kavejo/TransportAgents/wiki/HeaderAgent): Insert a custom header with a custom value in the header of every message that traverse the mail server
- [HeaderReroutingAgent](https://github.com/kavejo/TransportAgents/wiki/HeaderReroutingAgent): Rewrite the routing domain to be a differnt one (can be used to re-route traffic via a specific send connector matching the domain name space) when a specific header is present in the message
- [NDRAgent](https://github.com/kavejo/TransportAgents/wiki/NDRAgent): Drops Non-Delivery Report (NDR) that contains in the body the word "DELETE"
- [RecipientDomainRewriteAgent](https://github.com/kavejo/TransportAgents/wiki/RecipientDomainRewriteAgent): For any message sent to an address whose domain part is "contoso.com", it redirect the message to the same recipient on domain "tailspin.com"
- [SenderDomainRewriteAgent](https://github.com/kavejo/TransportAgents/wiki/SenderDomainRewriteAgent): For any message received from an address whose domain part is "contoso.com", it changes the sending domain and make the message appear from "tailspin.com"
- [TaggingAgent](https://github.com/kavejo/TransportAgents/wiki/TaggingAgent): Redirect a message to somebody+text@domain to somebody@domain, implementing Plus Addressing

## Wiki

For detailed information refer to the avaialble [Wiki](https://github.com/kavejo/TransportAgents/wiki)
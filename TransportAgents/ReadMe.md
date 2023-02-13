This sample code shows some Trnsport Agents that can be installed on Exchange 2019 CU12 (the DLLs referenced are for this version).

For the Transport Agent that intercept the Non Delivery Report (NDR) that contains the word "DELETE" can be installed as follows:

Install-TransportAgent -Name NDRAgent -TransportAgentFactory "TransportAgents.NDRAgent" -AssemblyPath "F:\Transport Agents\TransportAgents.dll"

For the Transport Agent that handled Plus Addressing (redirect a message to somebody+text@domain to somebody@domain) can be installed as follows:

Install-TransportAgent -Name TaggingAgent -TransportAgentFactory "TransportAgents.TaggingAgent" -AssemblyPath "F:\Transport Agents\TransportAgents.dll"

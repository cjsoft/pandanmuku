
import json
import socket
so=socket.socket(socket.AF_INET,socket.SOCK_DGRAM)
so.bind(("",1225))
a={"command":"register"}
so.sendto(json.dumps(a).encode(),("127.0.0.1",2333))
i,o=so.recvfrom(1024)
print (i,o)
a=json.loads(i.decode())
a["command"]="login"
a["password"]=a["password"]
a["username"]=a["username"]
so.sendto(json.dumps(a).encode(),("127.0.0.1",2333))
i,o=so.recvfrom(1024)
print(i,o)
while True:
	i,o=so.recvfrom(1024)
	j=json.loads(i.decode())
	print(i,o,j["content"])
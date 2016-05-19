from wsgiref.simple_server import *
from wsgiref import *
from cgi import *
import time
import socket
import json
import sys
import os
import threading
import pdb
import random
#webserver_
class webserver(object):
	def __init__(self,userver,sitepath,port,successbinary,errorbinary):
		if(os.path.exists(sitepath) and os.path.isfile((sitepath + "/site.cjsx").replace("//","/"))):
			f=open((sitepath + "/site.cjsx").replace("//","/")  ,"rb")
			self.imagelist={"jpg","png","bmp","jpeg","ico"}
			self.port=port
			self.rawsite=f.read()
			self.errb=errorbinary
			self.sucb=successbinary
			self.u=userver
			self.sitepath=sitepath
		else:
			print os.path.exists(sitepath) ,os.path.isfile(sitepath)
	def start(self):
		self._makenstartserver()
	def _makenstartserver(self):
		def request_handler(env,resp,*ags):
			try:
				if(env["REQUEST_METHOD"]=="GET"):
					ext=env["PATH_INFO"].lower().split(".")
					if(ext[-1] in self.imagelist and len(ext)>0):

						rtype=[("Content-Type","image/" + ext[-1])]
						if (os.path.isfile((self.sitepath + env["PATH_INFO"].lower().replace("/","/")).replace("//","/"))):
							resp("200 OK",rtype)
						else:
							resp("404 not found",[("Content-Type","text/html")])
							return ["<h1>404 NOT FOUND (*/w/*)</h1>"]
						image=open((self.sitepath + env["PATH_INFO"].lower().replace("/","/")).replace("//","/"),"rb")
						print time.asctime(),"returned",(self.sitepath + env["PATH_INFO"].lower().replace("/","/")).replace("//","/")
						if "wsgi.file_wrapper" in env:
							return env["wsgi.file_wrapper"](image,1024)
						else:
							return iter(lambda: image.read(1024),"")
					else:
						rtype=[("Content-Type","text/html")]
						resp("200 OK",rtype)
						spt=env["PATH_INFO"].split("/")
						if(len(spt)==2 or len(spt)==3):
							return [self.rawsite]
						else:
							doc=open((self.sitepath+env["PATH_INFO"].lower().replace("/","/")).replace("//","/"),"r")
							if "wsgi.file_wrapper" in env:
								return env["wsgi.file_wrapper"](doc,1024)
							else:
								return iter(lambda: doc.read(1024),"")
				elif(env["REQUEST_METHOD"]=="POST"):
					rtype=[("Content-Type","text/html")]
					resp("200 OK",rtype)
					if not("wsgi.input" in env or env["PATH_INFO"]=="/"):
						return [self.rawsite+self.errb]
					spt=env["PATH_INFO"].split("/")
					ext=dict(zip(range(len(spt)),spt))
					postdata=env["wsgi.input"].read(int(env.get("CONTENT_LENGTH",0)))
					psd=parse_qs(postdata)
					#psd["content"]=psd["content"][0]
					temp=json.loads(json.dumps(psd))
					#h=dict(range(len(tem)))
					if(self.u.postpandanmuku(temp.get(u"content",[u""])[0],env.get(u"HTTP_USER_AGENT",u"UNKNOWN"),ext.get(1,""))==True):
						return [self.rawsite+self.sucb]
					else:
						return [self.rawsite+self.errb]
			except BaseException as e:
				print time.asctime(),"web error:",str(e)
		def mse():
			try:
				httpd=make_server("",self.port,request_handler)
				httpd.serve_forever()
			except BaseException as e:
				print "web interface error while establishing,",str(e)
		t=threading.Thread(target=mse)
		t.start()
#udpcilents_management
class udpserver(object):
	def __init__(self,ip,port,jobjuserinfo,regenabled):
		def _listen():
			try:
				while True:
					if(self._server_activated==True):
						databuffer,cilentaddr=self.udp.recvfrom(1024)
						temp=json.loads(databuffer)
						temp["from"]=cilentaddr
						self.jsonhandler(temp)
			except BaseException as e:
				print "Error at udpserver._listen()",str(e)
				_listen()
		try:
			super(udpserver, self).__init__()
			self.ip=ip
			self.port=port
			self.udp=socket.socket(socket.AF_INET,socket.SOCK_DGRAM)
			self._server_activated=False
			self._tlisten=threading.Thread(target=_listen,args=())
			self.rege=regenabled
			self._timerinterval=60
			self._timertarget=self._heartbeatchecker
			self._timerargs=()
			self.rawusers=jobjuserinfo
			self.users=dict()
			self.usernip=dict()
			self._hb=dict()
			self.ipuser=dict()
			for i in jobjuserinfo:
				self.users[(i,jobjuserinfo[i])]=""
			self.udp.bind((ip,port))
			self.starttimer()
			self._tlisten.start()
		except BaseException as e:
			print "Error at udpserver.__init__()",str(e)
	def jsonhandler(self,jobj):
		try:
			if(jobj["command"]=="login"):
				if("username" in jobj and "password" in jobj):
					user=(jobj["username"],jobj["password"])
					if(user in self.users):
						if(self.users[user]!=""):
							self._sendtextmessage(self._createerror("failed, because you are over loginned"),self.usernip[user[0]])
						self.users[user]=jobj["from"]
						self.usernip[user[0]]=jobj["from"]
						self.ipuser[jobj["from"]]=user
						self._hb[user]=int(time.time())
						print time.asctime(),"loginned",user
						self._sendtextmessage( self._createservermessage("ok","successfully loginned"),jobj["from"])
					else:
						print time.asctime(),"wrong username or password",user
						self._sendtextmessage(self._createerror("login failed, false username or password"),jobj["from"])
				else:
					print time.asctime(),"wrong username or password",user
					self._sendtextmessage(self._createerror("login failed, illegal parameter"),jobj["from"])
			elif(jobj["command"]=="logout"):
				if self.ipuser.get(jobj["from"],"")=="":
					return 0
				print time.asctime(),"logout",self.ipuser[jobj["from"]]
				self.users[self.ipuser[jobj["from"]]]=""
				self.ipuser.pop(jobj["from"])
				self.usernip.pop(self.ipuser[jobj["from"]][0])
				self._hb.pop(self.ipuser[jobj["from"]])
			elif(jobj["command"]=="heartbeat"):
				if self.users.get(self.ipuser[jobj["from"]],("e","e"))==("e","e"):
					self._sendtextmessage(self._createerror("you are unloginned"),jobj["from"])
					return 0
				print time.asctime(),"heartbeat",self.ipuser[jobj["from"]]
				self._hb[self.ipuser[jobj["from"]]]=int(time.time())
			elif(jobj["command"]=="register"):
				if (self.rege== False):
					self._sendtextmessage(self._createerror("failed, registrations aren't allowed on this pandanmuku server"),jobj["from"])
				else:
					temp=self.rstr(5)
					pw=self.rstr(8)
					while(temp in self.rawusers):
						temp=self.rstr()
					self.rawusers[temp]=pw
					f=open("/storage/emulated/0/CJSoft/userjson.cjsx","w")
					f.write(json.dumps(self.rawusers))
					f.close()
					self.users[(temp,pw)]=""
					print time.asctime(),"registered",(temp,pw)
					self._sendtextmessage(self._createservermessage("ok","successfully registered",("username",temp),("password",pw)),jobj["from"])
		except BaseException as e:
			if jobj.get("command","")!="logout":
				print time.asctime(),"illegal data, command :",jobj.get("command","NULL"),"-details:",str(e)
				self._sendtextmessage(self._createerror("server error, perhaps you are unregistered, details:%s"%str(e)),jobj["from"])
	def _heartbeatchecker(self):
		try:
			poplist=list()
			for i in self.ipuser:
				if(int(time.time())-self._hb.get(self.ipuser[i],99999999)>60):
					poplist.append((i,self.ipuser[i]))
			for i in poplist:
				print time.asctime(),"kicked",i
				self.users[i[1]]=""
				self.usernip.pop(i[1][0])
				self.ipuser.pop(i[0])
				self._hb.pop(i[1])
				#self._sendtextmessage(self._createerror("kicked, because you have failed to send heartbeat pack frequently."),i[0])
		except KeyError:
			pass
	def _sendtextmessage(self,res,addr):
		if (self._server_activated):
			self.udp.sendto(res,addr)
	def _createerror(self,errinfo):
		rtn={}
		rtn["type"]="error"
		rtn["tag"]="[%s]"%time.asctime() + errinfo
		return json.dumps(rtn)
	def _createservermessage(self,typ,info="",*parameter):
		rtn=dict()
		rtn["type"]=typ
		rtn["tag"]="[%s]"%time.asctime() + str(info.encode("utf-8"))
		for i in parameter:
			rtn[i[0]]=i[1]
		return json.dumps(rtn)
	
	@property
	def timerinterval(self):
	    return self._timerinterval
	@timerinterval.setter
	def timerinterval(self, value):
		self._timerinterval = value

	def starttimer(self):
		def timer():
			while True:
				self._heartbeatchecker()
		t=threading.Thread(target=timer,args=())
		t.start()
		

	
	def start_server(self):
		self._server_activated=True

	def rstr(self,length):
		rtns= random.sample("1234567890qwertyuiopasdfghjklzxcvbnm",length)
		rtn=""
		return rtn.join(rtns)
	def postpandanmuku(self,content,useragent,usern):
		if usern in self.usernip:
			self._sendtextmessage (self._createservermessage("pandanmuku",content,("content",content),("useragent",useragent)),self.usernip[usern])
			return True
		else:
			return False
tempj="{}"
if(os.path.isfile("/storage/emulated/0/CJSoft/userjson.cjsx")):
	f=open("/storage/emulated/0/CJSoft/userjson.cjsx","r")
	tempj=f.read()
	f.close()
else:
	f=open("/storage/emulated/0/CJSoft/userjson.cjsx","w")
	f.close()
if(tempj==""):
	jo=dict()
else:
	jo=json.loads(tempj)
ud=udpserver("0.0.0.0",2333,jo,True)
ud.start_server()
web=webserver(ud,"/storage/emulated/0/CJSoft",8000,"<h1>Success</h1>","<h1>Failed</h1>")
web.start()

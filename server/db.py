def jsonhandler(self,jobj):
	try:
		if(jobj["command"]=="login"):
			if("username" in jobj["args"] and "password" in jobj["args"]):
				user=(jobj["args"]["username"],jobj["args"]["password"])
				if(user in self.users):
					if(self.users[user]!=""):
						self._sendtextmessage(self._createerror("failed, because you are over loginned"))
						return 0
					self.users[user]=jobj["from"]
					self.usernip[user[0]]=jobj["from"]
					self.ipuser[jobj["from"]]=user
					self._sendtextmessage( self._createservermessage("ok","successfully login"),jobj["from"])
			else:
				self._sendtextmessage(self._createerror("login failed, illegal parameter"),jobj["from"])
		elif(jobj["command"]=="logout"):
			if self.ipuser.get(jobj["from"],"")=="":
				self._sendtextmessage(self._createerror("you are unloginned"))
				return 0
			self.users[self.ipuser[jobj["from"]]]=""
			self.ipuser.pop(jobj["from"])
			self.usernip.pop(self.ipuser[jobj["from"]][0])
			self._hb.pop(self.ipuser[jobj["from"]])
			self._sendtextmessage( self._createservermessage("ok","successfully logout"),jobj["from"])
		elif(jobj["command"]=="heartbeat"):
			if (self.ipuser.get(jobj["from"],"")==""):
				self._sendtextmessage(self._createerror("you are unloginned"))
				return 0
			self._hb[self.users[jobj["from"]]]=True
		elif(jobj["command"]=="register"):
			if (self.rege== False):
				self._sendtextmessage(self._createerror("failed, registrations aren't allowed on this pandanmuku server"),jobj["from"])
			else:
				temp=self.rstr(5)
				pw=self.rstr(8)
				while(temp in self.rawusers):
					temp=self.rstr()
				self.rawusers[temp]=pw
				self.users[(temp,pw)]=""
				self._sendtextmessage(self._createservermessage("ok","successfully registered",("username",temp),("password",pw)))
	except BaseException:
		print (e)
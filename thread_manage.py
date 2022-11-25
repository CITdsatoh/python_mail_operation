#encoding:utf-8

import threading
import os
import time
import subprocess

class ThreadManage:

   def __init__(self,cmd):
     
     self.__cmd=cmd
     self.__using_file_resource=True
     self.__no_error=True
     
   
   def exec_cmd(self):
     self.__using_file_resource=True
     self.__no_error=True
     another_thread=threading.Thread(target=self.execute_cmd_in_another_thread)
     another_thread.start()
     
         
   def execute_cmd_in_another_thread(self):  
     try:
       proc=subprocess.Popen(self.__cmd,shell=True)
       proc.wait()
     except Exception:
       proc.kill()
       self.__no_error=False
       self.__using_file_resource=False
     finally:
       self.__using_file_resource=False
   
   def is_using_file_resource(self):
      return self.__using_file_resource
   
   def has_no_error(self):
      return self.__no_error
   
   
   
   
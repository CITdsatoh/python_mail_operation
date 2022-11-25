#encoding:utf-8

import tkinter as tk
from tkinter import simpledialog

STATES=("削除済みへ移動","保存","完全削除")

class OneMailInfo:

   next_id=0
   def __init__(self,one_data):
     each_data=one_data.split(",")
     #自動的にID(列番号）をつける
     self.__data_id=self.__class__.next_id
     
     self.__class__.next_id += 1
     self.__mail_address=each_data[0]
     self.__sender_name=each_data[1]
     self.__time_str=each_data[2]
     self.__mail_num=int(each_data[3])
     self.__state=each_data[4]
     if self.__state not in STATES:
       self.__state=STATES[0]
     self.__check_row=CheckedData()
   
   def get_disp_values(self):
     check_row_str=self.__check_row.current_check_str
     return (check_row_str,self.__data_id+1,self.__mail_address,self.__sender_name,self.__mail_num,self.__state)
   
   def change_check_state(self):
     self.__check_row.change_check()
     return self.__check_row.current_check_str
   
   def enable_check(self):
     self.__check_row.enable_check()
     return self.__check_row.current_check_str
   
   def disable_check(self):
     self.__check_row.disable_check()
     return self.__check_row.current_check_str
   
   def re_set_state(self,state):
     if state in STATES:
       self.__state=state
   
   #変更結果をファイルに書き込む際の文字列
   def __str__(self):
     return "%s,%s,%s,%d,%s"%(self.__mail_address,self.__sender_name,self.__time_str,self.__mail_num,self.__state)
   
   def __repr__(self):
     return "%s(\"%s,%s,%s,%d,%s\")"%(self.__class__.__name__,self.__mail_address,self.__sender_name,self.__time_str,self.__mail_num,self.__state)
   
   @property
   def data_id(self):
     return  self.__data_id
     
   @property
   def mail_address(self):
     return self.__mail_address
   
   
   @property
   def sender_name(self):
     return self.__sender_name
     
   
   @property
   def time_str(self):
     return self.__time_str
   
   @property
   def mail_num(self):
     return self.__mail_num
   
   @property
   def state(self):
     return self.__state
   
   @property
   def row_check_str(self):
     return self.__check_row.current_check_str
   
   @property
   def current_row_checked(self):
     return self.__check_row.current_check
   
   @classmethod
   def reset_id(cls):
     cls.next_id=0


class CheckedData:
 
   pseduo_check_str={"uncheck":"☐", "checked":"☑"}
   
   def __init__(self):
     self.__current_check=False
     self.__current_check_str=self.__class__. pseduo_check_str["uncheck"]
   
   def change_check(self):
     self.__current_check=not(self.__current_check)
     current_state=self.get_current_check_str()
     self.__current_check_str=self.__class__. pseduo_check_str[current_state]
   
   
   def enable_check(self):
     self.__current_check=True
     self.__current_check_str=self.__class__. pseduo_check_str["checked"]
   
   def disable_check(self):
     self.__current_check=False
     self.__current_check_str=self.__class__. pseduo_check_str["uncheck"]
   
   def get_current_check_str(self):
     if self.__current_check:
       return "checked"
     
     return "uncheck"
   
   
   @property
   def current_check_str(self):
     return self.__current_check_str
   
   @property
   def current_check(self):
     return self.__current_check
     

class SelectButtonDialog(simpledialog.Dialog):
   
    def __init__(self,parent,mail_info,title=None):
    
      self._mail_info=mail_info
      self._state_widget_var=tk.IntVar()
      #self._mov_choice=None
      #self._save_choice=None
      #self._del_choice=None
      super().__init__(parent,title)
      
    
    def body(self,body):
      return self
        
    def buttonbox(self):
      #box=tk.Frame(self)
      self.geometry("768x256")
      explain_text="宛先:"+self._mail_info.sender_name+"(メールアドレス:"+self._mail_info.mail_address+")からのメールを今後どのように扱いますか?\n以下の3つから選択をし,「OK」ボタンを押してください\nそして,最初の画面に存在する「設定を保存」ボタンを押してください。\nこれで設定の変更が保存されます(下の「OK」ボタンを押しただけでは設定は反映されないので気を付けてください)"
      explaination_label=tk.Label(self,text=explain_text,font=("helvetica",10,"bold"))
      explaination_label.place(x=0,y=8)
      mov_choice=tk.Radiobutton(self,text=STATES[0],variable=self._state_widget_var,value=0)
      save_choice=tk.Radiobutton(self,text=STATES[1],variable=self._state_widget_var,value=1)
      del_choice=tk.Radiobutton(self,text=STATES[2],variable=self._state_widget_var,value=2)
      try:
        self._state_widget_var.set(STATES.index(self._mail_info.state))
      #「削除済みへ移動」,「保存」,「完全削除」以外の値がセルに入っていたら
      except ValueError:
        self._state_widget_var.set(0)
      mov_choice.place(x=0,y=96)
      save_choice.place(x=0,y=144)
      del_choice.place(x=0,y=192)
      ok_button=tk.Button(self, text="OK", width=10, command=self.ok, default=tk.ACTIVE)
      ok_button.place(x=0,y=224)
      cancel_button=tk.Button(self, text="Cancel", width=10, command=self.cancel)
      cancel_button.place(x=96,y=224)
      self.bind("<Return>", self.ok)
      self.bind("<Escape>", self.cancel)
      self.bind("<KeyPress>",self.choice_by_key)
      
    
    def apply(self,event=None):
      new_state_index=self._state_widget_var.get()
      new_state=STATES[new_state_index]
      self.result=new_state
    
    
    def choice_by_key(self,event):
      pressed_key=event.keysym
      if pressed_key == "Up" or pressed_key == "Down":
        current_state=self._state_widget_var.get()
        new_state=(current_state-1)%3 if pressed_key == "Up" else (current_state+1)%3
        self._state_widget_var.set(new_state)
    
    def __str__(self):
      new_state_index=self._state_widget_var.get()
      new_state=STATES[new_state_index]
      return "選択状態:"+new_state
   
    




def selectstatedialog(parent,one_data):
   dialog=SelectButtonDialog(parent,one_data)
   ret=dialog.result or ""
   
   return ret
      
     
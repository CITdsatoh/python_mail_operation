#encoding:utf-8

import tkinter as tk
from tkinter import simpledialog
import re
import fnmatch
import unicodedata

STATES=("削除済みへ移動","保存","完全削除")

class OneMailInfo:

   next_id=0
   def __init__(self,one_data):
     #csvのファイルをセルごとに区切る場合は,区切り文字として「,」が使われるが,
     #「"」（ダブルクオーテーション)で囲まれているときには「,」は区切り文字とみなさず、区切りたくない
     each_data=split_except_escape_char(one_data,",","\"")
     
     
     #自動的にID(列番号）をつける
     self.__data_id=self.__class__.next_id
     
     self.__class__.next_id += 1
     
     self.__mail_address=each_data[0]
     self.__sender_name=each_data[1]
     self.__time_str=each_data[2]
     self.__cumulative_mail_num=int(each_data[3])
     self.__exists_mail_num=int(each_data[4])
     self.__receive_mail_num=int(each_data[5])
     self.__deleted_folder_mail_num=int(each_data[6])
     
     #実際にファイルに書き込むときの正式な状態
     self.__state=STATES[0]
     self.__display_only_state=STATES[0]
     if each_data[7] in STATES:
       self.__state=each_data[7]
       self.__display_only_state=each_data[7]
       
     self.__one_data_state_has_changed=True
     self.__check_row=CheckedData()
   
   def get_disp_values(self):
     check_row_str=self.__check_row.current_check_str
     mail_exists_num_disp_str="%d(受:%d,削済:%d)"%(self.__exists_mail_num,self.__receive_mail_num,self.__deleted_folder_mail_num)
     if self.__display_only_state != self.__state:
       return (check_row_str,self.__data_id+1,OneMailInfo.disp_remove_csv_escape(self.__mail_address),OneMailInfo.disp_remove_csv_escape(self.__sender_name),self.__cumulative_mail_num,mail_exists_num_disp_str,self.__tmp_new_state+"(未反映)")
       
     return (check_row_str,self.__data_id+1,OneMailInfo.disp_remove_csv_escape(self.__mail_address),OneMailInfo.disp_remove_csv_escape(self.__sender_name),self.__cumulative_mail_num,mail_exists_num_disp_str,self.__state)
   
   #こちらはテーブルに表示されている状態は新しくするが,
   #実際のファイルに書き込む際の正式な状態変更はまだ昔のままにするための一時変更メソッド
   def re_set_display_state(self,tmp_new_state):
     if tmp_new_state in STATES:
       self.__display_only_state=tmp_new_state
   
   #テーブルに表示されている新しい状態を昔の状態に戻す
   def cancel_renew_state(self):
     self.__display_only_state=self.__state
   
   def change_check_state(self):
     self.__check_row.change_check()
     return self.__check_row.current_check_str
   
   def enable_check(self):
     self.__check_row.enable_check()
     return self.__check_row.current_check_str
   
   def disable_check(self):
     self.__check_row.disable_check()
     return self.__check_row.current_check_str
   
   #表示されている状態を正式に反映する
   def renew_state(self):
     self.__state=self.__display_only_state
   
   
   #変更結果をファイルに書き込む際の文字列
   def __str__(self):
     return "%s,%s,%s,%d,%d,%d,%d,%s"%(self.__mail_address,self.__sender_name,self.__time_str,self.__cumulative_mail_num,self.__exists_mail_num,self.__receive_mail_num,self.__deleted_folder_mail_num,self.__state)
   
   def __repr__(self):
     return "%s(\"%s,%s,%s,%d,%d,%d,%d,%s\")"%(self.__class__.__name__,self.__mail_address,self.__sender_name,self.__time_str,self.__cumulative_mail_num,self.__exists_mail_num,self.__receive_mail_num,self.__deleted_folder_mail_num,self.__state)
   
   @property
   def data_id(self):
     return  self.__data_id
     
   @property
   def mail_address(self):
     return self.__mail_address
     
   
   @property
   def escaped_mail_address(self):
     return self.__class__.disp_remove_csv_escape(self.__mail_address)
   
   @property
   def sender_name(self):
     return self.__sender_name
   
   @property
   def escaped_sender_name(self):
     return self.__class__.disp_remove_csv_escape(self.__sender_name)
   
     
   
   @property
   def time_str(self):
     return self.__time_str
   
   @property
   def cumulative_mail_num(self):
     return self.__cumulative_mail_num
     
   @property
   def exists_mail_num(self):
     return self.__exists_mail_num
   
   @property
   def receive_mail_num(self):
     return self.__receive_mail_num
   
   @property
   def deleted_folder_mail_num(self):
     return self.__deleted_folder_mail_num
   
   @property
   def state(self):
     return self.__state
   
   @property
   def display_only_state(self):
     return self.__display_only_state
   
   @property
   def row_check_str(self):
     return self.__check_row.current_check_str
   
   @property
   def current_row_checked(self):
     return self.__check_row.current_check
   
   
   @classmethod
   def reset_id(cls):
     cls.next_id=0

   #この呼び出し元ファイルはCSVである
   #CSVでのダブルクオーテーションは,「ダブルクオーテーションで囲んだ範囲にあるカンマを区切り文字ではなくカンマそのものとして扱わせる」という目印
   #だが,CSVでダブルクオーテーションを文字列として扱いたいときは,""(ダブルクオーテーション)2つ並べる。このプログラムはCSV関係ないので,このまま呼び出すと,ダブルクオーテーションが2つ並んでしまう
   #よって、ここでダブルクオーテーション2つをダブルクオーテーション1つとして表示させる
   #また,ダブルクオーテーションが1つ,つまり,カンマをエスケープするという意味でのダブルクオーテーションは,ここではCSV関係ないので表示から外す。そのため,ダブルクオーテーション1つの時はダブルクオーテーションを見えなくする
   #いわば、逆エスケープ処理を行う
   #ただし,このテーブル処理終了後,内容は,またCSVとして文字列として保存しなおすので、インスタンス変数の内容自体を逆エスケープするわけではない。ユーザーに見せる表示部分のみを逆エスケープ
   @classmethod
   def disp_remove_csv_escape(cls,csv_str):
     removed_str=csv_str
     if "," in csv_str and csv_str.startswith("\"") and csv_str.endswith("\""):
       removed_str=csv_str[1:len(removed_str)-1]
      
     removed_str=removed_str.replace("\"\"","\"")
     
     return removed_str 



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
      explain_text="宛先:"+OneMailInfo.disp_remove_csv_escape(self._mail_info.sender_name)+"(メールアドレス:"+OneMailInfo.disp_remove_csv_escape(self._mail_info.mail_address)+")からのメールを今後どのように扱いますか?\n以下の3つから選択をし,「OK」ボタンを押してください\nそして,最初の画面に存在する「設定を保存」ボタンを押してください。\nこれで設定の変更が保存されます(下の「OK」ボタンを押しただけでは設定は反映されないので気を付けてください)"
      explaination_label=tk.Label(self,text=explain_text,font=("helvetica",10,"bold"))
      explaination_label.place(x=0,y=8)
      mov_choice=tk.Radiobutton(self,text=STATES[0],variable=self._state_widget_var,value=0)
      save_choice=tk.Radiobutton(self,text=STATES[1],variable=self._state_widget_var,value=1)
      del_choice=tk.Radiobutton(self,text=STATES[2],variable=self._state_widget_var,value=2)
      try:
        self._state_widget_var.set(STATES.index(self._mail_info.display_only_state))
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


#第一引数の文字列を第二引数の文字列を区切り文字として区切ってリストとして返すが,
#第三引数として与えた文字列に囲まれている間にある第二引数の文字列は区切り文字としてみなさないようにする
#不定長の先読みが使えないためこれで代用
#使用例としてcsvのファイルをセルごとに区切る場合は,区切り文字として「,」が使われるが,「"」（ダブルクオーテーション)で囲まれているときの「,」は区切り文字とみなさない
def split_except_escape_char(one_str:str,split_char,escape_char=""):
   result=[]
   split_start=0
   is_inside_escape_char=False
   
   for i in range(0,len(one_str)):
     if one_str[i] == split_char:
        if not is_inside_escape_char:
           result.append(one_str[split_start:i])
           split_start=i+1
           continue
     elif one_str[i] == escape_char:
        is_inside_escape_char=not is_inside_escape_char
   else:
     result.append(one_str[split_start:len(one_str)])
  
   return result
      
     
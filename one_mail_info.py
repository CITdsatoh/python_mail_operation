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
     each_data=one_data.split(",")
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
   
   def is_according_on_conditions(self,conditions:dict,type:str):
      if type == "pattern":
         return self.is_according_on_pattern_conditions(conditions)
      elif type == "num":
         return self.is_according_on_number_conditions(conditions)
      return True
       
   def is_according_on_pattern_conditions(self,conditions:dict):
      f_pattern=conditions["pattern"]
      f_basement=self.__sender_name if conditions["basement"] == "name" else self.__mail_address
      #メールアドレスや宛名側の空白を取り除く指定があった場合,ここで取り除く
      if conditions["remove_space"]:
         f_basement=re.sub("\\s+","",f_basement)
      for one_expr in conditions["expressions"]:
        #「肯定」検索の時は,このメソッド(入力したパターンに文字列があっているか)が1つでもTrueを返せばTrue
        #一方「否定」検索の時は,「肯定」検索のときにTrueが返されるパターンを満たしてはいけないので,この「肯定」検索の時のパターンしか返さないメソッドが1つでもTrueを返せばFalseになる
        if self.__class__.comp_expression_and_pattern(f_pattern,one_expr,f_basement,conditions["ignore_case"],conditions["ignore_char_width"]):
           #f_patternがnから始まらないものは「肯定」検索しているのでTrue,nから始まってしまうものは「否定」検索しているので,False
           return not f_pattern.startswith("n")
      
      #逆に１つも満たさなかった場合,「肯定」の時はFalse,「否定」の時はTrue
      return f_pattern.startswith("n")
     
   
   def is_according_on_number_conditions(self,conditions:dict):
     pattern=conditions["mail_num_pattern"]
     start=conditions["start"]
     end=conditions["end"]
     comp_num=self.__cumulative_mail_num
     if pattern == "exists" :
        comp_num=self.__exists_mail_num
     elif pattern == "receive":
        comp_num=self.__receive_mail_num
     elif pattern == "delete":
        comp_num=self.__deleted_folder_mail_num
     
     is_larger_than_start=(start < comp_num) if conditions["start_pattern"] == "gt" else (start <= comp_num)
     is_less_than_end=(comp_num < end) if conditions["end_pattern"] == "lt" else (comp_num <= end)
     
     #下限なし、上限なしの時は代わりに-1が入っているので,そのときは件数を満たしている扱いする
     return (is_larger_than_start or start == -1) and (is_less_than_end or end == -1)
   
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
   def sender_name(self):
     return self.__sender_name
     
   
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
   
   #ここでは,第一引数:与えられた比較パターン(前方一致:f,後方一致:b,部分一致:p,完全一致:e,ワイルドカード:wc,正規表現:re)
   #ごとに,第三引数(探索文字列)が第二引数(探索範囲)（を含んでいる|から始まる|で終わる|と一致する|表現を満たす)かを検索する
   @classmethod
   def comp_expression_and_pattern(cls,pattern_name,expr_pattern,text,is_ignore_case,is_ignore_char_width):
     #第一引数の比較パターンを表す文字列の最初にnが含まれていたらそれは否定(部分一致否定,前方一致否定・・etc)となるがそれは別メソッドで場合分けを行う
     #なのでここでは肯定の時だけを考える
     
     if is_ignore_char_width:
       expr_pattern=unicodedata.normalize("NFKC",expr_pattern)
       text=unicodedata.normalize("NFKC",text)
       
     if is_ignore_case:
      #大文字小文字を無視するワイルドカード・正規表現の比較は比較手法が特殊なので先にやってしまう
      if "wc" in pattern_name:
        return fnmatch.fnmatch(text,expr_pattern)
      if "re" in pattern_name:
        return re.search(expr_pattern,text,re.IGNORECASE) is not None
      expr_pattern=expr_pattern.lower()
      text=text.lower()
   
     #第一引数が部分一致(p)なら,第二引数を第三引数が含んでいればTrue
     if "p" in pattern_name:
       return expr_pattern in text
     #第一引数が前方一致(f)なら,第三引数が第二引数で始まればTrue
     if "f" in pattern_name:
       return text.startswith(expr_pattern)
     #第一引数が後方一致(b)なら,第三引数が第二引数で終わればTrue
     if "b" in pattern_name:
       return text.endswith(expr_pattern)
     #第一引数が完全一致(e)なら,第二引数と第三引数が等しければTrue
     #ただしeを含むとした場合,"re"(正規表現）の場合も入ってしまうのでそれは除く
     if "e" in pattern_name and "r" not in pattern_name:
       return expr_pattern == text
     
     #以下は大文字小文字を区別する比較
     if "wc" in pattern_name:
       return fnmatch.fnmatchcase(text,expr_pattern)
     if "re" in pattern_name:
       return re.search(expr_pattern,text) is not None

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
      
     
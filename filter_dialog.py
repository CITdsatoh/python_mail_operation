#encoding:utf-8

from tkinter import simpledialog
import tkinter as tk
import tkinter.ttk as ttk
import re
from filter_objs import PatternMatchingFilter,MailNumFilter


class FilterMakeDialog(simpledialog.Dialog):

   def __init__(self,parent,title=None):
   
     self.__filter_guide_strs={}
     self.__filter_guide_strs["p"]="のいずれかを含んでいる"
     self.__filter_guide_strs["f"]="のいずれかから始まる"
     self.__filter_guide_strs["b"]="のいずれかで終わる"
     self.__filter_guide_strs["e"]="の中のいずれかである"
     self.__filter_guide_strs["wc"]="のいずれかのパターンに一致する"
     self.__filter_guide_strs["re"]="のいずれかのパターンに一致する"
     
     self.__filter_guide_strs["np"]="のいずれも含まれない"
     self.__filter_guide_strs["nf"]="のいずれからも始まらない"
     self.__filter_guide_strs["nb"]="のいずれでも終わらない"
     self.__filter_guide_strs["ne"]="の中のいずれにも一致しない"
     self.__filter_guide_strs["nwc"]="のいずれのパターンに一致しない"
     self.__filter_guide_strs["nre"]="のいずれのパターンに一致しない"
    
     super().__init__(parent,title)
   
   def buttonbox(self):
   
     self.geometry("1056x608")
     self.__header_label=tk.Label(self,text="絞り込み表示\n以下にて選択・入力したテキストとパターンに合致するメールアドレスまたは宛名の列にのみ\n表示を絞り込みます",font=("times",18,"bold"))
       
     self.__pos_neg_radio_var=tk.IntVar()
     self.__positive_pattern_radio=tk.Radiobutton(self,text="肯定",variable=self.__pos_neg_radio_var,value=1,font=("times",14),command=self.guide_change)
     self.__negative_pattern_radio=tk.Radiobutton(self,text="否定",variable=self.__pos_neg_radio_var,value=0,font=("times",14),command=self.guide_change)
     
     
     self.__filter_pattern_var=tk.StringVar()
     self.__filter_select=ttk.Combobox(self,textvariable=self.__filter_pattern_var,height=6,state="readonly",values=("部分一致","前方一致","後方一致","完全一致","ワイルドカード","正規表現"))
     self.__filter_select.bind("<<ComboboxSelected>>",self.guide_change)
     
     self.__guide_label=tk.Label(self,text="以下の入力欄に絞り込んで表示させたい条件を入力してください。,（カンマ)で区切った場合,複数の条件を指定できます\nなお,複数指定した場合,いずれか1つ以上にあてはまるものに絞り込みます\nまた、,(カンマ）それ自体を含むものを検索したい場合はダブルクオーテーションで囲って,「\",\"」というようにしてください",font=("times",12,"bold"))
     
     self.__filter_obj_var=tk.StringVar()
     self.__filter_obj_select=ttk.Combobox(self,textvariable=self.__filter_obj_var,height=2,state="readonly",values=("メールアドレス","宛名"))
     self.__filter_obj_select.bind("<<ComboboxSelected>>",self.add_filter_basement_select_focus)
     
     
     
     self.__filter_guide_label=tk.Label(self,text="が",font=("times",14))
     self.__pattern_entry=tk.Entry(self,width=80,font=("times",14))
     self.__pattern_entry.bind("<Button-1>",self.unbind_keyevent)
     self.__entry_guide_label=tk.Label(self,text=self.__filter_guide_strs["p"],font=("times",12))
     
     
     self.__ignore_case_checkbox_var=tk.BooleanVar()
     self.__ignore_case_checkbox=tk.Checkbutton(self,text="大文字小文字を区別しない",variable=self.__ignore_case_checkbox_var,font=("times",12))
     
     self.__ignore_space_checkbox_var=tk.BooleanVar()
     self.__ignore_space_checkbox=tk.Checkbutton(self,text="メールアドレスの方の文字列に空白があるものは取り除いて検索する",variable=self.__ignore_space_checkbox_var,font=("times",12))
     
     self.__ignore_char_width_checkbox_var=tk.BooleanVar()
     self.__ignore_char_width_checkbox=tk.Checkbutton(self,text="半角全角を区別しない",variable=self.__ignore_char_width_checkbox_var,font=("times",12))
     
     self.__warning_label=tk.Label(self,text="",font=("helvetica",12,"bold"),fg="#ff0000")
     
     
     
     self.__header_label.place(x=4,y=4)
     self.__pos_neg_radio_var.set(1)
     self.__positive_pattern_radio.place(x=32,y=128)
     self.__negative_pattern_radio.place(x=128,y=128)
     
     
     self.__filter_select.current(0)
     self.__filter_select.place(x=32,y=192)
     self.__guide_label.place(x=16,y=224)
     
     self.__filter_obj_select.current(0)
     self.__filter_obj_select.place(x=16,y=312)
     self.__filter_guide_label.place(x=160,y=312)
     self.__pattern_entry.place(x=192,y=312)
     self.__entry_guide_label.place(x=160,y=344)
     
     self.__ignore_case_checkbox_var.set(False)
     self.__ignore_case_checkbox.place(x=256,y=372)
     
     self.__ignore_space_checkbox_var.set(False)
     self.__ignore_space_checkbox.place(x=256,y=400)
     
     self.__ignore_char_width_checkbox_var.set(False)
     self.__ignore_char_width_checkbox.place(x=256,y=432)
     
     self.__warning_label.place(x=4,y=464)
     
     self.__reset_btn=tk.Button(self,text="設定をキャンセル")
     self.__reset_btn.place(x=256,y=528)
     self.__reset_btn.bind("<Button-1>",self.reset_all_settings)
     self.__ok_btn=tk.Button(self,text="OK",width=10, command=self.ok, default=tk.ACTIVE)
     self.__ok_btn.place(x=400,y=528)
     self.__cancel_btn=tk.Button(self,text="Cancel", width=10, command=self.cancel)
     self.__cancel_btn.place(x=512,y=528)
     
     #ドロップダウンリストは2種類あるので今どちらのドロップダウンリストにフォーカスを与えるかを示す変数
     #0はパターン選択（前方一致,後方一致など),1はフィルター対象選択(メールアドレス,宛名)
     self.__current_vertical_focus=0
     
     self.bind("<Return>", self.ok)
     self.bind("<Escape>", self.cancel)
   
   def box(self):
     return  self.__filter_select
   
   
   def apply(self):
     
       #何をもとにしてフィルターをかけるかを表す(nameは宛名,mail_addressはメールアドレス)
       filter_base="name" if self.__filter_obj_select.get() == "宛名" else "mail_address"
       
       
       input_patterns=self.get_filter_patterns()
       
       if len(input_patterns) != 0:
        is_ignore_case=self.__ignore_case_checkbox_var.get()
        is_ignore_space=self.__ignore_space_checkbox_var.get()
        is_ignore_char_width=self.__ignore_char_width_checkbox_var.get()
        self.result=PatternMatchingFilter(filter_base,self.get_current_filter_pattern(),input_patterns,is_ignore_case,is_ignore_space,is_ignore_char_width)
   
   def operation_by_key_horizontal(self,event):
     pressed_key=event.keysym
     if pressed_key == "Left" or pressed_key == "Right":
         new_state=(self.__pos_neg_radio_var.get()+1)%2
         self.__pos_neg_radio_var.set(new_state)
         current_pattern=self.get_current_filter_pattern()
         self.__entry_guide_label["text"]=self.__filter_guide_strs[current_pattern]
   
   def operation_by_key_vertical(self,event):
     pressed_key=event.keysym
     if pressed_key == "Up" or pressed_key == "Down":
         current_select=self.__filter_select if self.__current_vertical_focus == 0 else self.__filter_obj_select
         value_num=len(current_select["values"])
         current_chosen_index=current_select.current()
         next_index=(current_chosen_index-1)%value_num if pressed_key == "Up" else (current_chosen_index+1)%value_num
         current_select.current(next_index)
         current_pattern=self.get_current_filter_pattern()
         self.__entry_guide_label["text"]=self.__filter_guide_strs[current_pattern]
         
   
   def add_filter_basement_select_focus(self,event=None):
     self.__current_vertical_focus=1
     self.__ignore_space_checkbox["text"]=self.__filter_obj_select.get()+"の方の文字列に空白があるものは取り除いて検索する"
     self.bind("<KeyPress>",self.operation_by_key_vertical)
   
   def reset_all_settings(self,event=None):
     self.__pos_neg_radio_var.set(1)
     self.__filter_select.current(0)
     self.__filter_obj_select.current(0)
     self.__ignore_case_checkbox_var.set(False)
     self.__ignore_space_checkbox_var.set(False)
     self.__ignore_char_width_checkbox_var.set(False)
     self.__pattern_entry.delete(0,tk.END)
     self.__entry_guide_label["text"]=self.__filter_guide_strs["p"]
     self.__warning_label["text"]=""
     self.__ignore_space_checkbox["text"]="メールアドレスの方の文字列に空白があるものは取り除いて検索する"
     self.__filter_select.focus_set()
     self.__current_vertical_focus=0
     
     
   def guide_change(self,event=None):
     if event is None:
       self.bind("<KeyPress>",self.operation_by_key_horizontal)
       self.bind("<KeyPress>",self.operation_by_key_horizontal)
     else:
       self.__current_vertical_focus=0
       self.bind("<KeyPress>",self.operation_by_key_vertical)
     
     current_pattern=self.get_current_filter_pattern()
     self.__entry_guide_label["text"]=self.__filter_guide_strs[current_pattern]
     if current_pattern == "ne":
       warn_str="注:現在「完全一致」の「否定」が選択されていますが,これは,入力されたテキストの始めから終わりまでまるっきり同一であるもの以外,\nつまり，1文字でもどこか一致するようなもの（部分一致など)は表示されます\nもしも,入力したテキストのうち1文字も一致しないものを検索したい場合は「部分一致」の「否定」で検索してみてください"
       self.__warning_label["text"]=warn_str
     elif current_pattern == "np":
       warn_str="注:現在「部分一致」の「否定」が選択されていますが,これは,入力されたテキストを1文字たりとも含まないもののみが表示されます.\nもしも,完全一致はしないけれども,どこかしら1文字以上は一致するものを検索する場合は「完全一致」の「否定」で検索してみてください"
       self.__warning_label["text"]=warn_str
     else:
       self.__warning_label["text"]=""
       
     
   #入力欄が触れられたとき,キーイベントのリスナーを解除するためのメソッド(疑似実装)
   #つまり,入力欄が触れられたときに何もしないメソッドを紐づける
   def unbind_keyevent(self,event):
     self.bind("<KeyPress>",self.key_operation_dummy)
   
   def key_operation_dummy(self,event):
      pass
     
   
   def get_current_filter_pattern(self):
     pattern=""
     if self.__pos_neg_radio_var.get() == 0:
        pattern="n"
     
     chosen_filter_way=self.__filter_select.get()
     if chosen_filter_way == "前方一致":
        return pattern+"f"
     if chosen_filter_way == "後方一致":
        return pattern+"b"
     if chosen_filter_way == "完全一致":
        return pattern+"e"
     if chosen_filter_way == "ワイルドカード":
        return pattern+"wc"
     if chosen_filter_way == "正規表現":
        return pattern+"re"
     
     return pattern+"p"
   
   def get_filter_patterns(self):
   
      #まずは,""(クォーテーション)でエスケープされていないカンマでパターンを区切る)
      input_patterns=re.split("(?<!\"),(?!\")",self.__pattern_entry.get())
      
      #空ではない（何らかの入力があった場合）
      if len(input_patterns) != 0:
       #次にクォーテーションでエスケープされたカンマを普通のカンマに直す(空文字列は除く）
       input_patterns=[one_pattern.replace("\",\"",",") for one_pattern in input_patterns if len(one_pattern) != 0]
      
      return input_patterns
     
   
   def __str__(self):
      return "hoge"

class MailNumFilterMakeDialog(simpledialog.Dialog):

   
   LABEL_NAMES=("累積(含完全削除済)の","現存する(両フォルダ)","受信フォルダの","削除済みフォルダの")
   ITEM_NAMES=("cumulative","exists","receive","delete")
   
   def __init__(self,parent,title=None):
     
     super().__init__(parent,title)
   
   def buttonbox(self):
     self.geometry("1024x512")
     self.__main_label=tk.Label(self,text="メールの数で絞り込み\n累積メール数が入力された数の範囲内であるものに絞り込みを行います\n左の入力欄は下限（以上）,右には上限(以下)を入力してください.\nなお,入力欄に何も入力しなかった場合は,メール数の下限(上限)はなしとみなします\nまた,ある特定の数きっかりの宛先を表示する場合,左右両入力欄に同じ数を入力して下さい.\n例えば,メール数がちょうど10件である宛先を見たい場合は両入力欄に10を入れてください",font=("helvetica",16,"bold"))
     self.__main_label.place(x=8,y=8)
     
     self.__pattern_choose_widget_var=tk.IntVar()
     self.__cumulative_radiobutton=tk.Radiobutton(self,text="累積メール数(含:完全削除済み)",variable=self.__pattern_choose_widget_var,command=self.change_guide_text_label,value=0,font=("times",16))
     self.__cumulative_radiobutton.place(x=64,y=160)
     self.__exists_radiobutton=tk.Radiobutton(self,text="現存するメール数(受信フォルダ+削除済みフォルダ)",variable=self.__pattern_choose_widget_var,command=self.change_guide_text_label,value=1,font=("times",16))
     self.__exists_radiobutton.place(x=448,y=160)
     self.__receive_radiobutton=tk.Radiobutton(self,text="受信フォルダのみのメール数",variable=self.__pattern_choose_widget_var,command=self.change_guide_text_label,value=2,font=("times",16))
     self.__receive_radiobutton.place(x=64,y=224)
     self.__delete_radiobutton=tk.Radiobutton(self,text="削除済みフォルダのみのメール数",variable=self.__pattern_choose_widget_var,command=self.change_guide_text_label,value=3,font=("times",16))
     self.__delete_radiobutton.place(x=448,y=224)
     self.__pattern_choose_widget_var.set(0)
     
     self.__guide_label=tk.Label(self,text="累積メール数が",font=("times",16))
     self.__guide_label.place(x=16,y=288)
     self.__start_entry=tk.Entry(self,width=6,font=("times",16))
     self.__start_entry.place(x=240,y=288)
     self.__start_num_label=tk.Label(self,text="件",font=("times",16))
     self.__start_num_label.place(x=320,y=288)
     self.__start_include_select=ttk.Combobox(self,height=2,values=("以上(を含む)","より多い(を含まない)"),font=("times",16),width=12,state="readonly")
     self.__start_include_select.place(x=352,y=288)
     self.__end_entry=tk.Entry(self,width=6,font=("times",16))
     self.__end_entry.place(x=528,y=288)
     self.__end_num_label=tk.Label(self,text="件",font=("times",16))
     self.__end_num_label.place(x=604,y=288)
     self.__end_include_select=ttk.Combobox(self,height=2,values=("以下(を含む)","未満(を含まない)"),font=("times",16),width=12,state="readonly")
     self.__end_include_select.place(x=640,y=288)
     
     self.__last_guide_label=tk.Label(self,text="件のメールに絞り込む",font=("times",16))
     self.__last_guide_label.place(x=128,y=320)
     
     self.__warning_label=tk.Label(self,text="",font=("helvetica",14,"bold"),fg="#ff0000")
     self.__warning_label.place(x=128,y=352)
     
     self.__just_btn=tk.Button(self,text="入力された件数きっかりにする",font=("times",16))
     self.__just_btn.place(x=64,y=432)
     self.__just_btn.bind("<Button-1>",self.just_padding)
     self.__reset_btn=tk.Button(self,text="設定をキャンセル",font=("times",16))
     self.__reset_btn.place(x=400,y=432)
     self.__reset_btn.bind("<Button-1>",self.reset_all_settings)
     self.__ok_btn=tk.Button(self,text="OK",width=10, command=self.ok, default=tk.ACTIVE,font=("times",16))
     self.__ok_btn.place(x=608,y=432)
     self.__cancel_btn=tk.Button(self,text="Cancel", width=10, command=self.cancel,font=("times",16))
     self.__cancel_btn.place(x=744,y=432)
     self.__start_include_select.current(0)
     self.__end_include_select.current(0)
     
     self.bind("<Return>", self.ok)
     self.bind("<Escape>", self.cancel)
     self.bind("<KeyPress>",self.operation_rbtns_by_key)
   
   def box(self):
     return self
   
   def validate(self):
    start=self.__start_entry.get()
    end=self.__end_entry.get()
    start_pattern="ge"
    end_pattern="le"
    
    try:
      #入力欄が空でなかった場合,intに変換する
      start=start and int(start)
      end=end and int(end)
    except ValueError:
      self.__warning_label["text"]="数値ではないものが入力されています.入力は数値のみにしてください!"
      return False
    else:
      start_pattern_num=self.__start_include_select.current()
      end_pattern_num=self.__end_include_select.current()
      if type(start) == int and  type(end) == int:
        if end < start:
           self.__warning_label["text"]="下限の値が上限の値より大きいです.必ず下限の値は上限の値より同じか小さくしてください!"
           return False
           
        if (start_pattern_num == 1 or end_pattern_num == 1) and (start == end):
           self.__warning_label["text"]="この場合,合致するものが1件も見つかりません\nもし,きっかり%d件のメールを表示させたい場合は,下限のほうは,「以上(を含む)」にし,\nなおかつ,上限のほうは,「以下(を含む)」にしてください!"%(start)
           return False
          
      #負数判定は,片方だけが入力されたときも行いたいのでここで行う
      #int型としか比較できないのであらかじめいちいちint型かどうかを出す
      if (type(start) == int and start < 0) or (type(end) == int and end < 0):
          self.__warning_label["text"]="負の数が入力されています.\n必ず検索する際は左側(下限)と右側(上限)は0以上の正の数を入力してください!"
          return False
          
      #ここからは正常終了時
      if start_pattern_num == 1:
         start_pattern="gt"
      if end_pattern_num == 1:
         end_pattern="lt"
      
      #空文字列(上限下限なし）の時は下限上限を表す数としては-1を数値として入れておく(これについては別に処理する)
      #もし,ここではlen関数を利用して空文字列であることを判定した場合,必ずしもstart,endという変数がstr型であるとは限らないのでエラーが出るかもしれない
      #なので,type関数でstrであるかどうかを判定することで,空文字列か否かを判定する(ここの時点では,int型の変数または,str型の空文字列しか来ないため,strなら空文字列と判断できるから)
      #またor演算子を使わない理由は空文字列だけではなく数値の0が入力されたときもFalseと判定されることによって,0という数値が入力された時も右辺の数(-1)が入ってしまうということを防ぐため
      if type(start) == str:
        start=-1
      if type(end) == str:
        end=-1
      
      mail_num_pattern=self.__class__.ITEM_NAMES[self.__pattern_choose_widget_var.get()]
      
      self.result=MailNumFilter(mail_num_pattern,start_pattern,end_pattern,start,end)
      
      return True
    
    return True
  
   def reset_all_settings(self,event): 
     self.__pattern_choose_widget_var.set(0)
     self.__start_include_select.current(0)
     self.__end_include_select.current(0)
     self.__start_entry.delete(0,tk.END)
     self.__end_entry.delete(0,tk.END)   
     self.__warning_label["text"]=""
     self.__guide_label["text"]=self.__class__.LABEL_NAMES[0]
   
   def just_padding(self,event):
     self.__start_include_select.current(0)
     self.__end_include_select.current(0)
     self.__warning_label["text"]=""
     start_num_str=self.__start_entry.get()
     end_num_str=self.__end_entry.get()
     #ここでは文字列を数値に変換したデータは直接利用しないが,実際に左右の入力欄に入力されている数字をそろえるために,ここで入力されたデータが数値かどうかを調べている
     try:
       #空文字列以外を変換
       start_num_str and int(start_num_str)
       end_num_str  and int(end_num_str)
     except ValueError:
       self.__warning_label["text"]="数値以外の値が入力されています.必ず数値を入力してください!"
     else:
       #空文字列でなかったら元の数はそのままで,空文字列だったほうには反対側の入力欄に入力されている数値を入れてあげる
       #or演算子を用いることでどちらに数値が入力されているかをいちいち場合分けせずに記述することができる
       start_num_str=start_num_str or end_num_str
       end_num_str=end_num_str or start_num_str
       self.__start_entry.delete(0,tk.END)
       self.__end_entry.delete(0,tk.END)
       self.__start_entry.insert(0,start_num_str)
       self.__end_entry.insert(0,end_num_str)   
     
   
   def operation_rbtns_by_key(self,event):
     pressed_key=event.keysym
     if pressed_key == "Left" or pressed_key == "Right":
        current_chosen_rbtn_index=self.__pattern_choose_widget_var.get()
        next_chosen_rbtn_index=(current_chosen_rbtn_index-1)%4 if pressed_key == "Left" else (current_chosen_rbtn_index+1)%4 
        self.__pattern_choose_widget_var.set(next_chosen_rbtn_index)
        self.change_guide_text_label()
     elif pressed_key == "Up" or pressed_key == "Down":
        current_chosen_rbtn_index=self.__pattern_choose_widget_var.get()
        up_or_down=int(current_chosen_rbtn_index/2)
        left_or_right=current_chosen_rbtn_index%2
        next_chosen_rbtn_index=((up_or_down-1)%2)*2+left_or_right if pressed_key == "Up" else ((up_or_down+1)%2)*2+left_or_right
        self.__pattern_choose_widget_var.set(next_chosen_rbtn_index)
        self.change_guide_text_label()
        
    
   
   def change_guide_text_label(self):
      self.__guide_label["text"]=self.__class__.LABEL_NAMES[self.__pattern_choose_widget_var.get()]
     
  

def askfilterpattern(parent):
   
   fd=FilterMakeDialog(parent,"絞り込み条件の指定")
   return fd.result

def askfiltermailnum(parent):
   nfd=MailNumFilterMakeDialog(parent,"累積のメール数での絞り込み条件の指定")
   return nfd.result
 
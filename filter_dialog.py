#encoding:utf-8

from tkinter import simpledialog
import tkinter as tk
import tkinter.ttk as ttk
import re


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
     self.__header_label=tk.Label(self,text="絞り込み表示\nメール情報を以下に入力したパターンに合致するものに絞り込みます",font=("times",18,"bold"))
       
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
        self.result={"basement":filter_base,"pattern":self.get_current_filter_pattern(),"expressions":input_patterns,"ignore_case":self.__ignore_case_checkbox_var.get(),"remove_space":self.__ignore_space_checkbox_var.get(),"ignore_char_width":self.__ignore_char_width_checkbox_var.get()}
     
   
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
     self.__ignore_space_checkbox.set(False)
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


def askfilterpattern(parent):
   
   fd=FilterMakeDialog(parent,"絞り込み条件の指定")
   return fd.result
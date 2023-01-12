#encoding:utf-8

import tkinter as tk
import tkinter.ttk as ttk
import os
import re
from tkinter import messagebox
from data_table import DataTable
from thread_manage import ThreadManage
from filter_dialog import askfilterpattern,askfiltermailnum

class MailDataApplication(tk.Tk):
 
   def __init__(self):
     super().__init__()
     self.geometry("1536x1024")
     self.title("outlookメール取り扱い")
     
     self.__title_label=tk.Label(self,text="outlookメール宛先一覧\nここではメールの宛先ごとにoutlook上でのメールを「保存」するか,「完全削除」(復元不可)するか,「削除済みへ移動」（復元可能）をするかを選択し\n、設定することができます.メールの取り扱いを変更したい行をクリックして選択してください.\nまた,表の左側のチェックをつけて下のボタンを押すことで複数一括で設定できます",font=("Helvetica",15))
     self.__title_label.place(x=32,y=8)
     
     mail_csv_file=self.__class__.get_desktop_dir()+"\\outlook_mail_dest_list.csv"
     mail_csv_backup_file=self.__class__.get_roming_dir()+"\\outlook_mail_dest_list.csv"
     
     self.__filter_btn=tk.Button(self,text="宛先またはメールアドレスをもとにフィルターする",font=("times",11))
     self.__filter_btn.bind("<Button-1>",self.table_filter)
     self.__filter_btn.place(x=512,y=128)
     self.__num_filter_btn=tk.Button(self,text="メールの件数をもとにフィルターする",font=("times",11))
     self.__num_filter_btn.bind("<Button-1>",self.table_filter)
     self.__num_filter_btn.place(x=896,y=128)
     self.__filter_remove_btn=tk.Button(self,text="フィルター解除",font=("times",11))
     self.__filter_remove_btn.bind("<Button-1>",self.table_remove_filter)
     self.__filter_remove_btn.place(x=1200,y=128)
     
     self.__data_table=DataTable(self,mail_csv_file,mail_csv_backup_file)
     self.__data_table.place(x=16,y=160)
     
     
     self.__sub_explaination_label=tk.Label(self,text="以下のボタンを押すと,チェックマークを付けた宛先に対して一括で「保存」か「完全削除」か「削除済みへ移動」かを自動で設定しなおし（選択を変更し）ます。\nただし,こちらも選択のみで,保存は「設定を保存」ボタンを押してください.",font=("times",12,"bold"))
     self.__sub_explaination_label.place(x=32,y=480)
     self.__all_check_button=tk.Button(self,text="すべてチェックを入れる",font=("times",11))
     self.__all_check_button.bind("<Button-1>",self.change_check)
     self.__all_check_button.place(x=224,y=560)
     self.__all_de_check_button=tk.Button(self,text="すべてチェックを外す",font=("times",11))
     self.__all_de_check_button.bind("<Button-1>",self.change_check)
     self.__all_de_check_button.place(x=448,y=560)
     
     self.__open_original_file_button=tk.Button(self,text="元のcsvファイルを開く",font=("times",11))
     self.__open_original_file_button.bind("<Button-1>",self.open_original_file)
     self.__open_original_file_button.place(x=672,y=560)
     self.__outlook_exe_button=tk.Button(self,text="outlookを操作する",font=("times",11))
     self.__outlook_exe_button.bind("<Button-1>",self.exe_vbs)
     self.__outlook_exe_button.place(x=896,y=560)
     
     self.__mail_save_button=tk.Button(self,text="チェックしたした宛先からのメールを「保存」するようにする",font=("times",10))
     self.__mail_save_button.bind("<Button-1>",self.set_state)
     self.__mail_save_button.place(x=32,y=608)
     self.__mail_delete_button=tk.Button(self,text="チェックした宛先からのメールを「完全削除」するようにする",font=("times",10))
     self.__mail_delete_button.bind("<Button-1>",self.set_state)
     self.__mail_delete_button.place(x=448,y=608)
     self.__mail_move_button=tk.Button(self,text="チェックした宛先からのメールを「削除済みへ移動」するようにする",font=("times",10))
     self.__mail_move_button.bind("<Button-1>",self.set_state)
     self.__mail_move_button.place(x=864,y=608)
     
     self.__set_save_button=tk.Button(self,text="設定を保存",font=("times",12))
     self.__set_save_button.bind("<Button-1>",self.save_file)
     self.__set_save_button.place(x=384,y=656)
     self.__exit_button=tk.Button(self,text="終了",font=("times",12))
     self.__exit_button.bind("<Button-1>",self.exit)
     self.__exit_button.place(x=736,y=656)
     self.__all_button_enable=True
     self.protocol("WM_DELETE_WINDOW", self.exit)
     self.__data_table.mainloop()
     self.mainloop()
     
     
  
   
   def button_state_change(self):
     self.__all_button_enable=not(self.__all_button_enable)
     state_str="normal" if self.__all_button_enable else "disable"
     self.__filter_btn["state"]=state_str
     self.__filter_remove_btn["state"]=state_str
     self.__open_original_file_button["state"]=state_str
     self.__outlook_exe_button["state"]=state_str
     self.__set_save_button["state"]=state_str
     self.__exit_button["state"]=state_str
     self.__all_check_button["state"]=state_str
     self.__all_de_check_button["state"]=state_str
     self.__mail_save_button["state"]=state_str
     self.__mail_delete_button["state"]=state_str
     self.__mail_move_button["state"]=state_str
     
     
   
   def set_state(self,event):
     if self.__all_button_enable:
       button_str=event.widget["text"]
       para_start=button_str.index("「")
       para_end=button_str.index("」")
       new_state=button_str[para_start+1:para_end]
       self.__data_table.mail_state_multiples_choose(new_state)
    
   
   
   def save_file(self,event=None):
     if self.__all_button_enable:
       try:
         self.__data_table.fwrite()
       except PermissionError:
         messagebox.showerror("エラー",self.__data_table.read_file+"というファイルが閉じられていないか、読み取り専用になっている可能性があります。もう一度ファイルを閉じる等をしてやり直してください")
         return False
       except OSError:
         messagebox.showerror("","")
         return False
       else:
         messagebox.showinfo("設定の完了","設定が保存されました")
         return True
   
   def change_check(self,event):
     if self.__all_button_enable:
       button_str=event.widget["text"]
       start_index=re.search("(入れる|外す)$",button_str).start()
       new_state=button_str[start_index:]
       if new_state == "入れる":
          self.__data_table.all_enable_check()
       elif new_state == "外す":
          self.__data_table.all_disable_check()
   
   def table_filter(self,event):
     if self.__all_button_enable:
       conditions=None
       type="pattern"
       if "宛先またはメールアドレス" in event.widget["text"]:
         conditions=askfilterpattern(self.__data_table)
       elif "メールの件数" in event.widget["text"]:
         conditions=askfiltermailnum(self.__data_table)
         type="num"
         
         
       if conditions is not None:
          self.__data_table.filter_table_display_row(conditions,type)
   
   def table_remove_filter(self,event):
      if self.__all_button_enable:
        self.__data_table.filter_remove()
      
             
      
   def open_original_file(self,event):
     if self.__all_button_enable:
       try:
         
         #設定が未反映のまま終了ボタンか「×」ボタンが押されたとき,保存されていない旨をユーザーに示し、保存してから開くかどうかを聞く
         do_save=False
         
         #こちらは,ファイル保存時にエラーが発生していないかどうかを確認し,保存時にエラーが発生したら閉じないようにするための変数
         is_file_open_ok=True
         
         if self.__data_table.has_changed_unsaved:
            do_save=messagebox.askyesno("設定の保存","メールの取り扱いに関する設定がまだ選択しただけであって,設定は保存されてません。保存しますか?")
         
         if do_save:
            is_file_open_ok=self.save_file()
         
         self.button_state_change()
         
         if is_file_open_ok:
           self.update()
           
           ftime_before_renewing=os.path.getmtime(self.__data_table.read_file)
           
           th=ThreadManage([self.__data_table.read_file])
           th.exec_cmd()
           while th.is_using_file_resource() and th.has_no_error():
             self.update()
             
             
           ftime_after_renewing=os.path.getmtime(self.__data_table.read_file)
           
           if ftime_before_renewing !=  ftime_after_renewing:
             self.renew_table()
           elif not th.has_no_error():
             messagebox.showerror("エラー","ファイルが開けなかったか,操作中にエラーが発生しました")     
       finally:
         self.button_state_change()
         self.update()
    
    
   
   def exe_vbs(self,event):
    if self.__all_button_enable:
    
      try:
        self.update()
        
        #設定が未反映のまま終了ボタンか「×」ボタンが押されたとき,保存されていない旨をユーザーに示し、保存してから開くかどうかを聞く
        do_save=False
         
        #こちらは,ファイル保存時にエラーが発生していないかどうかを確認し,保存時にエラーが発生したら閉じないようにするための変数
        is_exe_ok=True
         
        if self.__data_table.has_changed_unsaved:
          do_save=messagebox.askyesno("設定の保存","メールの取り扱いに関する設定がまだ選択しただけであって,設定は保存されてません。保存しますか?")
         
        if do_save:
          is_exe_ok=self.save_file()
        
        self.button_state_change()
        
        if is_exe_ok:
          th=ThreadManage([os.path.abspath("..\\mail_operation.vbs")])
          th.exec_cmd()
          while th.is_using_file_resource() and th.has_no_error():
            self.update()
          
          if th.has_no_error():
             self.renew_table()
          else:
             messagebox.showerror("エラー","ファイルが開けなかったか,実行時にoutlookの方でエラーが発生しました")
        
      finally:
        self.button_state_change()
        self.update()
           
   def renew_table(self):
     self.__data_table.place_forget()
     self.__data_table.destroy()
     mail_csv_file=self.__class__.get_desktop_dir()+"\\outlook_mail_dest_list.csv"
     mail_csv_backup_file=self.__class__.get_roming_dir()+"\\outlook_mail_dest_list.csv"
     self.__data_table=DataTable(self,mail_csv_file,mail_csv_backup_file)
     self.__data_table.place(x=0,y=160)
     self.update() 
     
        
   def exit(self,event=None):
     if self.__all_button_enable:
       #設定が未反映のまま終了ボタンか「×」ボタンが押されたとき,保存されていない旨をユーザーに示し,保存するかどうかを聞く
       do_save=False
       
       #こちらは,ファイル保存時にエラーが発生していないかどうかを確認し,保存時にエラーが発生したら閉じないようにするための変数
       is_close_ok=True
       
       if self.__data_table.has_changed_unsaved:
          do_save=messagebox.askyesno("設定の保存","メールの取り扱いに関する設定がまだ選択しただけであって,設定は保存されてません。保存しますか?")
       
       if do_save:
          is_close_ok=self.save_file()
          
       if is_close_ok:
          self.destroy()
         
       
       
   @classmethod
   def get_desktop_dir(cls):
     home_dir=os.getenv('USERPROFILE')
     #デスクトップディレクトリを表す名前は環境によって異なる.(必ずしもDesktopであると限らない)
     #そのため"Desktop","desktop","デスクトップ"の3通りあると考えられるためすべて試す
     desktop_names=("Desktop","desktop","デスクトップ")
     for desktop_name in desktop_names:
       dsktop_path=os.path.join(home_dir,desktop_name)
       if os.path.exists(dsktop_path):
         return dsktop_path
     
     return None
   
   @classmethod
   def get_roming_dir(cls):
     home_dir=os.getenv('USERPROFILE')
     return os.path.join(home_dir,"AppData","Roaming")
   
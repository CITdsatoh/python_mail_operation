#encoding:utf-8

import tkinter as tk
import tkinter.ttk as ttk
import os
import re
import subprocess
from tkinter import messagebox,filedialog
from datetime import datetime
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
     
     #フィルター結果の絞り込み条件
     self.__filter_conditions=None
     
     self.__filter_btn=tk.Button(self,text="宛先またはメールアドレスをもとにフィルターする",font=("times",11))
     self.__filter_btn.bind("<Button-1>",self.table_filter)
     self.__filter_btn.place(x=512,y=112)
     self.__num_filter_btn=tk.Button(self,text="メールの件数をもとにフィルターする",font=("times",11))
     self.__num_filter_btn.bind("<Button-1>",self.table_filter)
     self.__num_filter_btn.place(x=896,y=112)
     self.__filter_remove_btn=tk.Button(self,text="フィルター解除",font=("times",11))
     self.__filter_remove_btn.bind("<Button-1>",self.table_remove_filter)
     self.__filter_remove_btn.place(x=1200,y=112)
     
     self.__data_table=DataTable(self,mail_csv_file,mail_csv_backup_file)
     self.__data_table.place(x=16,y=144)
     
     
     self.__sub_explaination_label=tk.Label(self,text="以下のボタンを押すと,チェックマークを付けた宛先に対して一括で「保存」か「完全削除」か「削除済みへ移動」かを自動で設定しなおし（選択を変更し）ます。\nただし,こちらも選択のみで,保存は「設定を保存」ボタンを押してください.",font=("times",12,"bold"))
     self.__sub_explaination_label.place(x=32,y=504)
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
     
     self.__set_save_button=tk.Button(self,text="設定を保存",font=("times",12))
     self.__set_save_button.bind("<Button-1>",self.save_file)
     self.__set_save_button.place(x=1120,y=560)
     
    
     self.__mail_save_button=tk.Button(self,text="チェックしたした宛先からのメールを「保存」するようにする",font=("times",10))
     self.__mail_save_button.bind("<Button-1>",self.set_state)
     self.__mail_save_button.place(x=32,y=608)
     self.__mail_delete_button=tk.Button(self,text="チェックした宛先からのメールを「完全削除」するようにする",font=("times",10))
     self.__mail_delete_button.bind("<Button-1>",self.set_state)
     self.__mail_delete_button.place(x=448,y=608)
     self.__mail_move_button=tk.Button(self,text="チェックした宛先からのメールを「削除済みへ移動」するようにする",font=("times",10))
     self.__mail_move_button.bind("<Button-1>",self.set_state)
     self.__mail_move_button.place(x=864,y=608)
     
    
     
     self.__restore_tmp_state_button=tk.Button(self,text="未反映の設定を変更前に戻す(設定の保存後はできません)",font=("times",12))
     self.__restore_tmp_state_button.bind("<Button-1>",self.restore_tmp_state)
     self.__restore_tmp_state_button.place(x=128,y=656)
     
     self.__only_filtered_save_button=tk.Button(self,text="フィルター後に表示されている宛先をファイルに書き込む",font=("times",12));
     self.__only_filtered_save_button.bind("<Button-1>",self.save_address_only_filtered);
     self.__only_filtered_save_button.place(x=640,y=656)
     
     self.__exit_button=tk.Button(self,text="終了",font=("times",12))
     self.__exit_button.bind("<Button-1>",self.exit)
     self.__exit_button.place(x=1120,y=656)
     
   
     self.__all_button_enable=True
     
     self.protocol("WM_DELETE_WINDOW", self.exit)
     self.__data_table.mainloop()
     self.mainloop()
     
     
  
   
   def button_state_change(self):
     self.__all_button_enable=not(self.__all_button_enable)
     state_str="normal" if self.__all_button_enable else "disable"
     self.__open_original_file_button["state"]=state_str
     self.__outlook_exe_button["state"]=state_str
     self.__set_save_button["state"]=state_str
     self.__exit_button["state"]=state_str
     self.__all_check_button["state"]=state_str
     self.__all_de_check_button["state"]=state_str
     self.__mail_save_button["state"]=state_str
     self.__mail_delete_button["state"]=state_str
     self.__mail_move_button["state"]=state_str
     self.__restore_tmp_state_button["state"]=state_str
     
   
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
     
     new_conditions=None
     if "宛先またはメールアドレス" in event.widget["text"]:
       new_conditions=askfilterpattern(self.__data_table)
     elif "メールの件数" in event.widget["text"]:
       new_conditions=askfiltermailnum(self.__data_table)
             
     if new_conditions is not None:
       self.__filter_conditions=new_conditions
       is_filter_succeed=self.__data_table.filter_table_display_row(self.__filter_conditions)
       if not is_filter_succeed:
          self.__filter_conditions=None
   
   def table_remove_filter(self,event):
     self.__filter_conditions=None
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
     self.__filter_conditions=None
     self.update() 
     
   
   def restore_tmp_state(self,event):
     if self.__all_button_enable:
       if self.__data_table.has_changed_unsaved:
          self.__data_table.cancel_changing_renew_state()   
       else:
          messagebox.showerror("エラー","ファイルに書き込んで保存され反映したものに関しては戻せません\n(戻せるのはメールの取り扱いの設定変更後、ファイルに書き込んでいない未反映のものに限ります)")
      
        
   #全宛先ではなく,フィルターした結果、絞り込まれた宛先のみをファイルに書き込む
   def save_address_only_filtered(self,event=None):
     
     if self.__all_button_enable:
       
       #フィルター結果の保存を行うかどうか(途中でユーザが保存するのを取り消すことができるようにするため,ユーザーが保存しようしているのかフラグとして残しておく)
       do_save=False
       
       #ユーザーが不正な入力などを行った際,ファイル名をね続けるかどうかのフラグ
       do_check=True
       
       #取り扱い設定の保存を行わず直接フィルターの保存をした場合は,そのまま示し
       guide_str=""
       
       if self.__data_table.has_changed_unsaved:
           #メールの取り扱い設定がすんでいなかった場合,最初に設定の保存をするかどうか否か聞く
          pre_do_save=messagebox.askyesno("設定の保存","フィルターを行う前に,まず,メールの取り扱いに関する設定がまだ選択しただけであって,設定は保存されてません。保存しますか?")
          if pre_do_save:
            do_check=self.save_file()
          #取り扱い設定の保存を行った場合は「次に」という文字列を挿入する
          guide_str="次に"
       
          
       #取り扱いを保存した後,何も前触れもなくフィルター結果の保存に移った場合,ユーザーが混乱するかもしれないので,その旨を示しておく
       if do_check:
         messagebox.showinfo("フィルター結果の保存",guide_str+"フィルター結果の保存に入ります。フィルター結果を保存したいファイルを選択してください")
             
          
       
       #保存先ファイル名(デフォルトでは,outlook_mail_filtered_dest_list_(現時刻のタイムスタンプ(14桁)).csvとする)
       now=datetime.now()
       save_csv_file_name="outlook_mail_filtered_dest_list_%04d%02d%02d%02d%02d%02d.csv"%(now.year,now.month,now.day,now.hour,now.minute,now.second)
       
       while do_check:
          chosen_file_name=filedialog.asksaveasfilename(title="フィルター結果の保存先ファイルを選択",initialdir=self.__class__.get_desktop_dir(),initialfile=save_csv_file_name,filetypes=[("csvファイル","*.csv")],defaultextension="csv")
          if len(chosen_file_name) == 0 or chosen_file_name == ".csv" :
             do_check=messagebox.askyesno("保存先ファイル選択の続行","ファイル名が入力されていませんが,もう一度保存先ファイルの選択を続行しますか?")
             continue
          if "outlook_mail_dest_list.csv" in chosen_file_name :
             messagebox.showerror("エラー",chosen_file_name+"には保存できません!別のファイルを選択してください!")
             do_check=messagebox.askyesno("保存先ファイル選択の続行","もう一度保存先ファイルの選択を続行しますか?")
             continue
          
          #正常だった場合(拡張子がcsvでなくても,何らかの入力があれば,正常扱いする)
          do_check=False
          do_save=True
          #拡張子がcsvでない場合はこちらで補う
          path,extension=os.path.splitext(chosen_file_name)
          if extension != ".csv":
             chosen_file_name += ".csv"
          save_csv_file_name=chosen_file_name
       
       
       if do_save:
          try:
             self.__data_table.filtered_write(save_csv_file_name)
             condition_file=save_csv_file_name.rstrip(".csv")+"_filter_condition.txt"
             with open(condition_file,"w",encoding="utf-8") as f:
                #self.__filter_conditionsは現在かかっているフィルターの内容を示すオブジェクト
                #これがNoneなら,フィルターがかかっていないので,代わりに"フィルター：なし"という文字列を書き込み,Noneでないとき(フィルターがあるときは),その内容を書く
                condition_write_str=self.__filter_conditions and str(self.__filter_conditions) or "フィルター:なし"
                f.write(condition_write_str+"\n")
                f.close()
          except (PermissionError,OSError):
             messagebox.showerror("ファイルの書き込み失敗","ファイルの書き込みに失敗しました。書き込もうとしているファイルが開かれているか,読み取り専用になっている可能性があります")
          else:
             messagebox.showinfo("ファイル書き込み完了","フィルター結果をファイルに書き込みました")
             save_csv_file_dir=save_csv_file_name[0:save_csv_file_name.rindex("/")]
             save_csv_file_dir_cmd_str="\""+save_csv_file_dir.replace("/","\\")+"\""
             subprocess.Popen(f"explorer {save_csv_file_dir_cmd_str}", shell=True)
           
   
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
        try:
          self.destroy()
        finally:
          self.__data_table.create_backup_file()
         
       
       
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
   
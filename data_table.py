#encoding:utf-8

import tkinter as tk
import tkinter.ttk as ttk
import os
import re
import stat
import shutil
import pathlib
import pickle
from datetime import datetime
from one_mail_info import OneMailInfo,selectstatedialog
from tkinter import messagebox

class DataTable(tk.Frame):

   def __init__(self,master,original_file,backup_file):
    
     self.__master=master
     super().__init__(self.__master,width=1312,height=384)
     self.__original_file=original_file
     self.__backup_file=backup_file
     self.__read_file=self.__original_file if os.path.exists(self.__original_file) else self.__backup_file
     self.__file_contents=[]
     
     one_file=None
     
     try:
        one_file=open(self.__read_file,encoding="utf_8_sig")
        self.__file_contents=one_file.readlines()
     #原本もバックアップもどちらもなかった場合
     except FileNotFoundError:
        pickle_file="table_data_backup.pickle"
        file_obj=open(pickle_file,"rb")
        pickled_backup_obj=pickle.load()
        file_obj.close()
        self.__file_contents=pickled_backup_obj.get_file_contents()
        pickled_backup_obj.restore_log_file()
     finally:
       one_file.close()
     
     self.__mail_info=[]
     
     self.__current_disp_info_ids=[]
     
     #状態(内容)変更後に保存が反映されていないかどうかを示すフラグ
     self.__has_changed_unsaved=False
     
     #こちらは状態変更を保存後、あるいは何も変更が行われていないにもかかわらず,アプリを終了するときにいちいちバックアップを作るのは面倒
     #ゆえに,内容の変更が保存された後,バックアップが作られたかどうかを示し,これがTrueならバックアップを作る
     self.__once_saved_must_backup=False
     
     for i in range(1,len(self.__file_contents)-1):
       self.__mail_info.append(OneMailInfo(self.__file_contents[i].strip("\n")))
       
     
     self.__col_names=("選択","No.","メールアドレス","宛名","累積メール数","現存メール数","メールの取り扱い")
     self.__col_widths=(48,48,240,240,272,272,192)
     self.__table=ttk.Treeview(self,columns=self.__col_names,height=15)
     
     #ctrlが押されているか,shiftが押されているかを表すフラグ
     self.__ctrl_pressing=False
     self.__shift_pressing=False
     
     #ctrlとshiftが同時に押されたときに実行される範囲選択の行番号を格納する
     #今は何も選択されていないのでNoneを格納しておく
     self.__chosen_range=[None,None]
     
     self.bind("<KeyPress>",self.judge_ctrl_shift_press)
     self.bind("<KeyRelease>",self.judge_ctrl_shift_release)
     self.__table.bind("<KeyPress>",self.judge_ctrl_shift_press)
     self.__table.bind("<KeyRelease>",self.judge_ctrl_shift_release)
     self.mk_table()
     self.__table.focus_set()
   
   def mk_table(self):
     
     
     self.__table.column("#0",width=0,anchor=tk.CENTER,stretch=False)
     
     #テーブルのの幅の設定
     for col_name,col_width in zip(self.__col_names,self.__col_widths):
      self.__table.column(col_name,width=col_width,anchor=tk.CENTER)
     
     headers=self.__class__.get_headers(self.__file_contents[0])
     
     self.__table.heading("#0",text="\n\n",anchor=tk.CENTER)
     #テーブルののヘッダーの設定
     for col_name,header in zip(self.__col_names,headers):
      self.__table.heading(col_name,text=header,anchor=tk.CENTER)
      
     
     
     
     #データの設定
     for table_id,one_mail_info in enumerate(self.__mail_info):
       self.__current_disp_info_ids.append(table_id)
       self.__table.insert("","end",iid=table_id,values=one_mail_info.get_disp_values())
     
     
     #フォントの設定
     font_style_header=ttk.Style()
     font_style_header.configure("Treeview.Heading",font=("Helvetica",10,"bold"),foreground="black")
     font_style_body=ttk.Style()
     font_style_body.configure("Treeview",font=("times",12))
     
     self.__table.bind("<Button-1>",self.table_operation) 
     self.__table.pack(side=tk.LEFT)
     y_scrollbar=ttk.Scrollbar(self,orient=tk.VERTICAL,command=self.__table.yview)
     self.__table.configure(yscroll=y_scrollbar.set)
     y_scrollbar.pack(side=tk.RIGHT,fill="y")
     
   
  
   def destroy(self):
     OneMailInfo.reset_id()
     super().destroy()
   
   #左クリックされたときに発動
   #単純な左クリックの場合は,単一の列(宛先)についてのみのメールの取り扱いを決める処理(ダイアログを出す)を行う
   #「左クリック+ctrlキー」なら複数選択の処理を行う
   def table_operation(self,event):
     try:
       table_id=int(self.__table.identify_row(event.y))
       #ctrlボタンが押されている
       if self.__ctrl_pressing:
         #両方とも押されているとき(範囲選択）
         if self.__shift_pressing:
            #選択範囲の行番号を格納(0番目は範囲の開始,1番目は範囲の終了)
            self.__chosen_range[0]=self.__chosen_range[1]
            self.__chosen_range[1]=table_id
            #0番目がNoneではない、つまり,範囲の開始と終了の両方が定まったとき実行(0番目がNoneならまだ開始の行しか定まっていないということ）
            if self.__chosen_range[0] is not None:
               start=min(self.__chosen_range[0],self.__chosen_range[1])
               end=max(self.__chosen_range[0],self.__chosen_range[1])
               
               #現在表に表示されているのは,フィルターがかかっていて,IDが0番のものから最後まで必ずしも連続で並んでいるわけではないので
               #現在,表示されている列のID番号を連続ではないが順番に格納しているcurrent_disp_info_idsを用いてインデックス情報を得る
               #選択範囲の開始の行と終了の行がcurrent_disp_info_idsのどのインデックスに入っているかを調べる
               start_index=self.__current_disp_info_ids.index(start)
               end_index=self.__current_disp_info_ids.index(end)+1
               
               #今選ばれたところがすべて選択済みなのかあるいは未選択なのかを得る
               #すでに選択済みのものばかりであれば,すべて選択を外し,すべて未選択であれば,選択済みにするが,
               #複数選択したところが,選択済みのものと未選択のものが混じっている場合は,現在の状態に関係なく全部選択済みにする
               #そのために現在の選択範囲がどうなっているかを調べている
               #なお,フィルターがかかっているかもしれないので,表示されているところ(current_disp_info_ids)に入っている列番号のみを参照する
               current_chosen=[self.__mail_info[index].current_row_checked for index in self.__current_disp_info_ids[start_index:end_index]]
               
               #配列current_disp_info_idsの開始の行が入っているインデックス番目から終了の行が入っているインデックス番号番目までの要素が現在の選択範囲となる
               for index in range(start_index,end_index):
                  #行番号
                  one_row_id=self.__current_disp_info_ids[index]
                  if all(current_chosen) or not any(current_chosen):
                    self.change_check(one_row_id)   
                    continue
                  check_str=self.__mail_info[one_row_id].enable_check()
                  self.__table.set(one_row_id,self.__col_names[0],check_str)
         #ctrlしか押されていないとき
         else:
            self.change_check(table_id)
            self.__chosen_range[1]=table_id
       #どっちも押されていないとき
       else:
         self.all_disable_check()
         self.__chosen_range=[None,None]
         self.mail_state_choose(table_id)
     #テーブルの列以外が触れられた時,identify_rowメソッドは数字以外の文字列を返すため,このエラーが発生する
     #その際は何もしなくてよいので例外を握りつぶす
     #何もキーを押していない状態で,テーブルの列以外が押された場合は,範囲選択を解除する
     #その際,テーブルの列以外が触れられた時は,特に何もしないので,握りつぶす
     except ValueError:
       pass
     finally:
       self.__table.selection_remove(self.__table.selection())
    
     
   #Ctrlキー、あるいはshiftが押されているかを判別するための必要
   #ctrlが押されたらTrueにする
   def judge_ctrl_shift_press(self,event):
     pressed_key=event.keysym
     if pressed_key == "Control_L":
       self.__ctrl_pressing=True
     if pressed_key == "Shift_L":
       self.__shift_pressing=True
     
   
   #ctrlからボタンが離されたらFalseにする
   def judge_ctrl_shift_release(self,event):
     release_key=event.keysym
     if release_key == "Control_L":
       self.__ctrl_pressing=False
     if release_key == "Shift_L":
       self.__shift_pressing=False
   
   
   def filter_table_display_row(self,conditions,type):
    
     for one_table_id in self.__current_disp_info_ids:
       self.__table.delete(one_table_id)
         
     self.__current_disp_info_ids=[]
     for table_id,one_mail_info in enumerate(self.__mail_info):
       if one_mail_info.is_according_on_conditions(conditions,type):
          self.__table.insert("","end",iid=table_id,values=one_mail_info.get_disp_values())
          self.__current_disp_info_ids.append(table_id)
       else:
          one_mail_info.disable_check()
          one_mail_info.cancel_renew_state()
     
     if len(self.__current_disp_info_ids) == 0:
        messagebox.showerror("エラー","条件に当てはまる項目が1つもありませんでした!")
        self.filter_remove()
     else:
        messagebox.showinfo("フィルター完了","%d件当てはまる項目が見つかりました"%(len(self.__current_disp_info_ids)))
             
     
   def filter_remove(self):
     for one_table_id in self.__current_disp_info_ids:
        self.__table.delete(one_table_id)
     self.__current_disp_info_ids=[]
     
     for table_id,one_mail_info in enumerate(self.__mail_info):
       self.__table.insert("","end",iid=table_id,values=one_mail_info.get_disp_values())
       self.__current_disp_info_ids.append(table_id)
        
   
   def mail_state_choose(self,table_id):
     
     new_state=selectstatedialog(self,self.__mail_info[table_id])
     if len(new_state) != 0:
      #状態変更したけど,未保存だよ
      self.__has_changed_unsaved=True
      self.__table.set(table_id,self.__col_names[6],new_state+"(未反映)")
      #メールの取り扱いについて,テーブルの表示は新しい状態にするが,実際の変更はまだされていないという状態にする
      self.__mail_info[table_id].re_set_display_state(new_state)
   
   def change_check(self,table_id):
     new_check_str=self.__mail_info[table_id].change_check_state()
     self.__table.set(table_id,self.__col_names[0], new_check_str)
   
   #一時的に行った状態変更（状態変更を行ったものの,まだファイルの保存がすんでいない未反映の状態)を元の状態に戻す
   #状態変更後,ファイルに保存する前まで仮の状態を元に戻す
   #元に戻すものがないときは発動されないよう別のクラスで処理しているので,ここで,戻すものが1つもなく発動されることはない
   def cancel_changing_renew_state(self):
     
     for table_id,one_mail_info in enumerate(self.__mail_info):
        if one_mail_info.display_only_state != one_mail_info.state:
          one_mail_info.cancel_renew_state()
          self.__table.set(table_id,self.__col_names[6],one_mail_info.state)
         
   
     #状態が元に戻ったので,変更でファイルに書き込んで反映されていないことを表すフラグを元に戻す
     self.__has_changed_unsaved=False
   
   def mail_state_multiples_choose(self,new_state):
     for table_id,one_mail_info in enumerate(self.__mail_info):
       if one_mail_info.current_row_checked:
         self.__has_changed_unsaved=True
         self.__mail_info[table_id].re_set_display_state(new_state)
         self.__table.set(table_id,self.__col_names[6],new_state+"(未反映)")
   
   #表示されているところだけチェックを入れる
   def all_enable_check(self):
     for table_id in self.__current_disp_info_ids:
       check_str=self.__mail_info[table_id].enable_check()
       self.__table.set(table_id,self.__col_names[0],check_str)
   
   #表示されているところだけチェックを外す
   def all_disable_check(self):
     for table_id in self.__current_disp_info_ids:
       check_str=self.__mail_info[table_id].disable_check()
       self.__table.set(table_id,self.__col_names[0],check_str)
     
     self.__table.selection_remove(self.__table.selection())
     
       
   def fwrite(self):
  
    
    with open(self.__original_file,"w",encoding="utf_8_sig") as f:
   
       #ヘッダは元ファイルのまま書き込む
       f.write(self.__file_contents[0])
       for table_id,one_mail_info in enumerate(self.__mail_info):
         one_mail_info.renew_state()
         f.write(str(self.__mail_info[table_id])+"\n")
         if table_id in self.__current_disp_info_ids:
           self.__table.set(table_id,self.__col_names[6],self.__mail_info[table_id].state)
       #フッター（合計）もそのまま書き込む
       f.write(self.__file_contents[-1])
       f.close()
    
    self.__table.selection_remove(self.__table.selection())  
     
    #バックアップファイルの方は安全上読み取り専用にしている。なのでここでいったん読み取り専用を解除する
    #だがバックアップファイル自体がない時はモード変更の仕様がないので,そのまま握りつぶす   
    try:
      os.chmod(path=self.__backup_file,mode=stat.S_IWRITE)
    except FileNotFoundError:
      pass
    
    try:
      shutil.copyfile(self.__original_file,self.__backup_file)
    except OSError:
       pass
    else:
      #再び読み取り専用に戻す
      os.chmod(path=self.__backup_file,mode=stat.S_IREAD)
    
    #こちらはファイルの保存は行われたので閉じる際にバックアップを作らなければいけないことを示すフラグ
    #これが立ったので,アプリ終了時にバックアップを作ることとなる
    self.__once_saved_must_backup=True
    
    
    #その後未反映ラベルを解除し,メールの取り扱い状態を変更する
    self.__has_changed_unsaved=False
    
   
   def create_backup_file(self):
    if self.__once_saved_must_backup:
     now_time=datetime.now()
     backup_log_dir=self.__backup_file[0:self.__backup_file.rindex("\\")+1]+"backup"
     if not pathlib.Path(backup_log_dir).exists():
        pathlib.Path(backup_log_dir).mkdir()
        
     now_time_str="%d%02d%02d%02d%02d%02d"%(now_time.year,now_time.month,now_time.day,now_time.hour,now_time.minute,now_time.second)
     backup_log_file=backup_log_dir+"\\outlook_mail_dest_list_"+now_time_str+"_backup.csv"
     try:
       shutil.copyfile(self.__original_file,backup_log_file)
     except OSError:
       pass
     else:
       os.chmod(path=backup_log_file,mode=stat.S_IREAD)
     
    #万が一の際のシリアライズ化用データ
    mail_data_contents=[]
    #ヘッダ
    mail_data_contents.append(self.__file_contents[0].strip("\n"))
     
    #ボディ(元データ)の付け加え
    mail_data_contents.extend(self.__mail_info)
     
    #フッター
    mail_data_contents.append(self.__file_contents[-1].strip("\n"))
     
    mail_log_file=self.__backup_file[0:self.__backup_file.rindex("\\")+1]+"datelog.log"
    backup_obj=SerializableBackUpInfoObj(tuple(mail_data_contents),mail_log_file)
    pickle_file="table_data_backup.pickle"
    with open(pickle_file,"wb") as f:
       pickle.dump(backup_obj,f)
       f.close()
      
    self.__once_saved_must_backup=False
     
   @property
   def read_file(self):
     return self.__read_file
   
   @property
   def has_changed_unsaved(self):
     return self.__has_changed_unsaved

    
   @classmethod
   def get_headers(cls,original_header):
     header_items=original_header.strip("\n").split(",")
     new_header=[]
     new_header.append("選択")
     new_header.append("No.")
     new_header.append(header_items[0])
     new_header.append(header_items[1])
     now_date=datetime.now()
     first_date="%04d/%02d/%02d"%(now_date.year,now_date.month,now_date.day)
     last_date=first_date
     match_obj=re.findall("[0-9]{4}/[0-1]?[0-9]/[0-3]?[0-9]",header_items[3])
     if len(match_obj) != 0:
       first_date=match_obj[0]
       last_date=match_obj[-1]
     
     new_header.append(first_date+"から"+last_date+"までの累積\n(現存しない完全削除済みも含む)")
     new_header.append(last_date+"時点での現存メール数\n(受信フォルダ+削除済みフォルダ)")
     new_header.append("この宛先のメールの取り扱い")
     
     return new_header


#何らかのアクシデントで元ファイルが消えてしまった場合に備え,終了するたびに毎回,ここまでのファイル情報とメールの最終確認時刻情報をシリアライズしておく
class SerializableBackUpInfoObj:
    
    def __init__(self,mails_data_info:tuple,fin_checked_data_log_file_path:str):
       self.__mails_data_info=mails_data_info
       self.__fin_checked_data_log_file_path=fin_checked_data_log_file_path
       date_info_file_obj=open(self.__fin_checked_data_log_file_path,encoding="utf_8_sig")
       date_info_contents=date_info_file_obj.readlines()
       date_info_file_obj.close()
       self.__fin_checked_data_str=date_info_contents[-1].strip("\n")
     
    #どの時刻分までのメールを確認したかを書いたlogファイルの復帰
    #こちらに関しては,存在の有無に関係なく,データファイル情報と同期していなければならないので,データを復帰した場合は必ずこちらは復帰をする
    def restore_log_file(self):
      try:
         os.chmod(path=self.__fin_checked_data_log_file_path,mode=stat.S_IWRITE)
      except FileNotFoundError:
         pass
      with open(self.__fin_checked_data_log_file_path,"w",encoding="utf_8_sig") as f:
         f.write(self.__fin_checked_data_str+"\n")
         f.close()
         
      os.chmod(path=self.__fin_checked_data_log_file_path,mode=stat.S_IREAD)
    
    #原本もバックアップも存在しない場合を前提とするので,存在確認や読み取り専用解除はここではしなくてよい
    def restore_data_files(self,original_file:str,backup_file:str):
       for one_file in (original_file,backup_file):
         with open(original_file,"w",encoding="utf_8_sig") as f:
           f.writelines(self.get_file_contents())
           f.close()
           
       #バックアップファイルの読み取り専用かだけはしておく  
       os.chmod(path=backup_file,mode=stat.S_IREAD)
    
    def get_file_contents(self):
      file_contents=[]
      for one_info in self.__mails_data_info:
         file_contents.append(str(one_info)+"\n")
      return file_contents
    
    #デバッグ用
    def __str__(self):
       log_file_data="log_file_path:"+self.__fin_checked_data_log_file_path+"\nfinal_mail_checked_time:"+self.__fin_checked_data_str+"\n"
       data=f"datas:{self.mail_info}\n"
       return log_file_data+data
    
    @property
    def mail_info(self):
      return list(self.__mails_data_info)[1:len(self.__mails_data_info)-1]
       
      
      
       
      
      
       
     
  
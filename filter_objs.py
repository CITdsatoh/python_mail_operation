#encoding:utf-8

import re
import fnmatch
import unicodedata

class PatternMatchingFilter:

   PATTERN_MATCHING_NAMES={"p":"部分一致","f":"前方一致","b":"後方一致","e":"完全一致","wc":"ワイルドカード","re":"正規表現"}
   BASEMENT_NAMES={"name":"宛名","mail_address":"メールアドレス"}
   PATTERN_MATCHING_JAPANESE={"p":"のいずれかがつく","f":"のいずれかから始まる","b":"のいずれかで終わる","e":"のいずれかに一致する","wc":"のいずれかのパターンにマッチする","re":"のいずれかの表現にマッチする","np":"のいずれもつかない","nf":"のいずれからでも始まらない","nb":"のいずれでも終わらない","ne":"のいずれにも一致しない","nwc":"のいずれのパターンにもマッチしない","nre":"のいずれの表現にもマッチしない"}
   
   def __init__(self,comp_basement:str,match_pattern:str,expressions,is_ignore_case:bool,is_remove_space:bool,is_ignore_char_width:bool):
      
      #比較基準(nameは宛名,mail_addressはメールアドレスで比較)
      self.__comp_basement=comp_basement
      #f:前方一致,b:後方一致,p:部分一致,e:完全一致,wcはワイルドカード,reは正規表現
      #nがつくとそれぞれの否定(例:nfは前方一致否定(～から始まらない))
      self.__match_pattern=match_pattern
      
      self.__expressions=expressions
      
      self.__is_ignore_case=is_ignore_case
      self.__is_remove_space=is_remove_space
      self.__is_ignore_char_width=is_ignore_char_width
   
   def is_according_on_conditions(self,mail_info):
      comp=mail_info.escaped_sender_name if self.__comp_basement == "name" else mail_info.escaped_mail_address
      #メールアドレスや宛名側の空白を取り除く指定があった場合,ここで取り除く
      if self.__is_remove_space:
          comp=re.sub("\\s+","",comp)
      for one_search in self.__expressions:
        #「肯定」検索の時は,このメソッド(入力したパターンに文字列があっているか)が1つでもTrueを返せばTrue
        #一方「否定」検索の時は,「肯定」検索のときにTrueが返されるパターンを満たしてはいけないので,この「肯定」検索の時のパターンしか返さないメソッドが1つでもTrueを返せばFalseになる
        if self.is_pattern_match(one_search,comp):
           #f_patternがnから始まらないものは「肯定」検索しているのでTrue,nから始まってしまうものは「否定」検索しているので,False
           return not self.__match_pattern.startswith("n")
      
      #逆に１つも満たさなかった場合,「肯定」の時はFalse,「否定」の時はTrue
      return self.__match_pattern.startswith("n")
   
   #ここでは,与えられた比較パターン(前方一致:f,後方一致:b,部分一致:p,完全一致:e,ワイルドカード:wc,正規表現:re)
   #ごとに,第二引数(探索文字列)が第一引数(探索範囲)（を含んでいる|から始まる|で終わる|と一致する|表現を満たす)かを検索する
   def is_pattern_match(self,search,text):
  
     #第一引数の比較パターンを表す文字列の最初にnが含まれていたらそれは否定(部分一致否定,前方一致否定・・etc)となるがそれは別メソッドで場合分けを行う
     #なのでここでは肯定の時だけを考える
     
     if self.__is_ignore_char_width:
       search=unicodedata.normalize("NFKC",search)
       text=unicodedata.normalize("NFKC",text)
       
     if self.__is_ignore_case:
      #大文字小文字を無視するワイルドカード・正規表現の比較は比較手法が特殊なので先にやってしまう
      if "wc" in self.__match_pattern:
        return fnmatch.fnmatch(text,search)
      if "re" in self.__match_pattern:
        return re.search(search,text,re.IGNORECASE) is not None
      search=search.lower()
      text=text.lower()
   
     #部分一致(p)なら,第一引数を第二引数が含んでいればTrue
     if "p" in self.__match_pattern:
       return search in text
     #前方一致(f)なら,第二引数が第一引数で始まればTrue
     if "f" in self.__match_pattern:
       return text.startswith(search)
     #後方一致(b)なら,第二引数が第一引数で終わればTrue
     if "b" in self.__match_pattern:
       return text.endswith(search)
     #完全一致(e)なら,第一引数と第二引数が等しければTrue
     #ただしeを含むとした場合,"re"(正規表現）の場合も入ってしまうのでそれは除く
     if "e" in self.__match_pattern and "r" not in self.__match_pattern:
       return search == text
     
     #以下は大文字小文字を区別する比較
     if "wc" in self.__match_pattern:
       return fnmatch.fnmatchcase(text,search)
     if "re" in self.__match_pattern:
       return re.search(search,text) is not None
   
   #文字列化(フィルター条件をファイルに書き込む際に必要)
   def __str__(self):
      header="フィルター内容"
      
      match_name=self.__match_pattern
      pos_neg="肯定"
      expr_sep="か"
      if self.__match_pattern.startswith("n"):
         pos_neg="否定"
         expr_sep="と"
         match_name=match_name.lstrip("n")
      
      pattern_match_str=f" 検索方法:{self.__class__.BASEMENT_NAMES[self.__comp_basement]}に対する{self.__class__.PATTERN_MATCHING_NAMES[match_name]}の{pos_neg}"
      
      expression_split_str=f"」{expr_sep}「"
      expression_str=expression_split_str.join(self.__expressions)
      expressions_str=f" 検索パターン:「{expression_str}」{self.__class__.PATTERN_MATCHING_JAPANESE[self.__match_pattern]}"
      
      is_ignore_case_str=" 大文字小文字違いの無視:あり" if self.__is_ignore_case else "大文字小文字違いの無視:なし"
      is_remove_space_str=" 空白の無視:あり" if self.__is_remove_space else "空白の無視:なし\n"
      is_ignore_char_width_str=" 半角全角の違いの無視:あり" if self.__is_ignore_char_width else "半角全角の違いの無視:なし"
      
      
      return "\n".join([header,pattern_match_str,expressions_str,is_ignore_case_str,is_remove_space_str,is_ignore_char_width_str])
      



class MailNumFilter:

    BASEMENT_NAMES={"cumulative":"累積(含完全削除済)の","exists":"現存する(両フォルダ)の","receive":"受信フォルダの","delete":"削除済みフォルダの"}
    LESS_PATTERN_STRS={"le":"以下","lt":"未満"}
    GREATER_PATTERN_STRS={"ge":"以上","gt":"より多い(を含まない)"}
    
    
    def __init__(self,comp_basement_folder_name:str,greater_pattern:str,less_pattern:str,range_min_mail_num:int,range_max_mail_num:int):
      
      #どのフォルダのメール数で比較するかの基準
      self.__comp_basement_folder_name=comp_basement_folder_name
      
      #最大値を含めるか,含めないか(geかgtが入る)
      self.__greater_pattern=greater_pattern
      
      #最大値を含めるか,含めないか(leかltが入る)
      self.__less_pattern=less_pattern
      
      #この変数は基本的に０以上の値しか入らないが,
      #-1の値が例外として入っている場合がある。その際はメール数の上限(下限)なしとみなす
      self.__range_min_mail_num=range_min_mail_num
      self.__range_max_mail_num=range_max_mail_num
    
    def is_according_on_conditions(self,mail_info):
      comp_folder_mail_num=mail_info.cumulative_mail_num
      if self.__comp_basement_folder_name == "exists":
         comp_folder_mail_num=mail_info.exists_mail_num
      elif self.__comp_basement_folder_name == "receive":
         comp_folder_mail_num=mail_info.receive_mail_num
      elif self.__comp_basement_folder_name == "delete":
         comp_folder_mail_num=mail_info.deleted_folder_mail_num
      
      is_greater_than_min=(self.__range_min_mail_num < comp_folder_mail_num) if self.__greater_pattern == "gt" else (self.__range_min_mail_num <= comp_folder_mail_num)
      is_less_than_max=(comp_folder_mail_num < self.__range_max_mail_num) if self.__less_pattern == "lt" else (comp_folder_mail_num <= self.__range_max_mail_num)
      
      #下限なし、上限なしの時は代わりに-1が入っているので,そのときは件数を満たしている扱いする
      return (is_greater_than_min or self.__range_min_mail_num == -1) and (is_less_than_max or self.__range_max_mail_num == -1)
    
    #文字列化(フィルター後の宛先を書き込む際にどんな条件でフィルターしたかを記録しておくために必要)
    def  __str__(self):  
      header_str="フィルター内容"
     
      basement_folder_num_str=" "+self.__class__.BASEMENT_NAMES[self.__comp_basement_folder_name]+"メールの件数でのフィルター"
      
      min_mail_num_str_header=" メール数の下限:"
      min_mail_num_str_body="%d件"%(self.__range_min_mail_num)+self.__class__.GREATER_PATTERN_STRS[self.__greater_pattern]
      if self.__range_min_mail_num == -1:
         min_mail_num_str_body="なし"
      
      max_mail_num_str_header=" メール数の上限:"
      max_mail_num_str_body="%d件"%(self.__range_max_mail_num)+self.__class__.LESS_PATTERN_STRS[self.__less_pattern]
      if self.__range_max_mail_num == -1:
         max_mail_num_str_body="なし"
      
      mail_num_str_full="".join([min_mail_num_str_header,min_mail_num_str_body,max_mail_num_str_header,max_mail_num_str_body])
      
      return "\n".join([header_str,basement_folder_num_str,mail_num_str_full])


class MailOperationStateFilter:
  
  def __init__(self,disp_tmp_delete_states:bool,disp_save_states:bool,disp_delete_states:bool,is_include_temporaily_changed:bool):
     self.__is_disp_dict_by_states={}
     self.__is_disp_dict_by_states["削除済みへ移動"]=disp_tmp_delete_states
     self.__is_disp_dict_by_states["保存"]=disp_save_states
     self.__is_disp_dict_by_states["完全削除"]=disp_delete_states
     
     #状態の変更はしたけれど,まだ正式に反映されていないものをどのように表示するかを示すフラグ
     #Trueなら,(まだ反映されていないが),変更後の状態
     #Falseなら,変更する前のもとの状態
     self.__is_include_temporaily_changed=is_include_temporaily_changed
  
  def is_according_on_conditions(self,mail_info):
    current_state=mail_info.display_only_state if self.__is_include_temporaily_changed else mail_info.state
    try:
      return self.__is_disp_dict_by_states[current_state]
    #誤って変な値(削除済みへ移動,保存,完全削除以外の値)がセットされていたら,辞書のkeyがないのでkeyError
    #その時のメールの取り扱いは「削除済みへ移動」と同じなので,変な値が入ったいるものの表示は,削除済みへ移動が表示されているかどうかに従うこととする
    except keyError:
      return self.__is_disp_dict_by_states["削除済みへ移動"]
  
  #フィルター後の状態をファイルに書き込む際のどのような条件でフィルターを行ったか書き残しておくためのもの
  def __str__(self):
     
     each_state_disp_info=[]
     for one_state,is_disp in self.__is_disp_dict_by_states.items():
       do_save_str="する" if is_disp else "しない"
       each_state_str=f" メールの取り扱いが「{one_state}」のもの:表示{do_save_str}"
       each_state_disp_info.append(each_state_str)
     
     all_info=[]
     all_info.append("フィルター内容")
     all_info.extend(each_state_disp_info)
     
     if self.__is_include_temporaily_changed:
       all_info.append(" 未反映の取り扱い:状態変更済みだが正式に反映されていないものも含む")
     else:
       all_info.append(" 未反映の取り扱い:状態変更後正式に反映済みのもののみ")
     
     
     return "\n".join(all_info)
      

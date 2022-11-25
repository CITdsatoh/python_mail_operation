#encoding:utf-8 

from mail_state_operation import MailDataApplication
import tkinter as tk
from tkinter import messagebox

if __name__ == "__main__":
  try:
    app=MailDataApplication() 
  except FileNotFoundError:
    sub=tk.Tk()
    sub.geometry("1024x1024")
    messagebox.showerror("メール情報","outlookのメール情報がありません")
    sub.destroy()
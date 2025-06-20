
import os
import sys
import tkinter as tk
from tkinter import filedialog
from reader import ExcelReader

def main():
    print("AFAD KK Excel Asistanı Mekana Girdi!")
    excel_reader = ExcelReader()

    while True:
        print("\nDile benden ne dilersen ('son' dersen biter, 'resize' excel boyuna ayar ceker, 'check' kontrol noktamız):")
        user_input = input().strip().lower()

        if user_input == "son":
            print("Yine bekleriz...")
            break
        elif user_input == "resize":
            excel_reader.resize_cells()
        elif user_input == "check":
            excel_reader.check_condition()
        else:
            print("Dile dedik ama tabi dile kolay. Seceneklerden secersen yaparım.")

if __name__ == "__main__":
    main()
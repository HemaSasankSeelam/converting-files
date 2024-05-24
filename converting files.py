
from tkinter import filedialog
import os,sys
from pathlib import Path
import pdf2docx
import docx2pdf
import img2pdf
import termcolor
from tkinter import *
import PyPDF2
import pdf_compressor
from PIL import Image
import requests
from tkinter import simpledialog
import pymsgbox
import cv2

root=Tk(className=" CONVERTING FILES".upper())
root.config(bg="#324145")
root.geometry("600x500")

class Converting_files:
    def pdf_docx(self,input_path=None,output_folder=None):
        try:
            password = None
            if output_folder is None:  # for single file 
                input_path = Path(filedialog.askopenfile(mode='r', filetypes=[("PDF FILES", '*.pdf')],
                                                        title="Select PDF File To Convert Into Word.").name)
                if self.check_is_encrypted(input_path):
                    password = simpledialog.askstring(prompt="Enter Password Of PDF: ", title="Selected Encrypted PDF.")
                    if password is None:
                        termcolor.cprint(text="\nEnter The Password Please Try Again....\n", color='red')
                        return
                file = pdf2docx.Converter(pdf_file=input_path, password=password)
                output_folder = filedialog.askdirectory(title="Select The Folder To Save Word File.")
                if output_folder == "" or output_folder is None:
                    termcolor.cprint(text="\nNo Output Folder Is Selected....\n", color='red')
                    return
                output_folder = Path(output_folder)
                output_path = output_folder / (input_path.stem + ".docx")
                file.convert(output_path)
                termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n", color='green')

            elif input_path != None and output_folder != None: # for multipul files

                li = [i.as_posix() for i in Path(input_path).glob("*.*pdf")]
                for i in li:
                    if self.check_is_encrypted(i):
                        termcolor.cprint(text = f"\nFile '{i}' Is Encrypted Decrypt The File To Convert Into 'docx' \n",color='red')
                        return
                for i in li:
                    file = pdf2docx.Converter(pdf_file=i,password=None)
                    output_path = Path(output_folder) / (Path(i).stem + ".docx")
                    file.convert(output_path)
                    print("\033[2J","\033[H")#clear the screen and return to the home position

                termcolor.cprint(text=f"\nFiles Is Saved On '{output_folder}' Folder\n",color='green')
        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected\n",color='red')
        except pdf2docx.converter.ConversionException:
            termcolor.cprint(text=f"\nEntered Wrong Password Check Once '{password}' \n",color='red')
        except Exception as e:
            print(e)
    
    def check_is_encrypted(self,input_path) -> bool:
        with open (input_path,'rb') as fo:
            pdf_reader = PyPDF2.PdfReader(fo)
            if pdf_reader.is_encrypted:
                return True
        return False
    
    def docx_pdf(self,input_path=None,output_folder=None):
        try:
            if output_folder == None: #Means for single files
                input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("DOCX FILES",'*.docx')],
                                                    title="Select Word File To Covert Into PDF.").name)
                output_folder = filedialog.askdirectory(title="Select Folder To Save The PDF File.")
                if output_folder == "" or output_folder == None:
                    termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                    return
                output_path = Path(output_folder) / (Path(input_path).stem + ".pdf")
                docx2pdf.convert(input_path,output_path)
                termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color="green")
            
            elif input_path == None and output_folder != None: # for multipul files
                li = [i.as_posix() for i in Path(input_path).glob("*.*docx")]
                for i in li:
                    output_path = Path(output_folder) / (Path(i).stem + ".pdf")
                    docx2pdf.convert(i,output_path)
                termcolor.cprint(text=f"\nFiles Is Saved On '{output_folder}' Folder\n",color='green')
        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)            
    
    def img_pdf(self):
        try:
            input_path = filedialog.askopenfile(mode='r',title="Select JPG/PNG File Only.",
                                                     filetypes=[("PNG FILES",'*.png'),("JPG Files","*.jpg")]).name
            output_folder = filedialog.askdirectory(title="Select The Folder To Save PDF.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            output_path = Path(output_folder) / (Path(input_path).stem + ".pdf")
            with open(output_path,'wb') as fo:
                fo.write(img2pdf.convert(input_path))
            termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color='green')

        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
   
    def jpeg_png(self):
        try:
            input_path=Path(filedialog.askopenfile(mode='r',filetypes=[("JPEG files","*.jpg")],
                                                   title="Select The JPEG File.").name)
            output_path = Path(input_path).parent / (Path(input_path).stem + ".png")
            Image.open(input_path).save(output_path)
            print(termcolor.colored(f"\nFile Is Saved On '{output_path}' \n",color="green"))
        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)

    def png_jpeg(self):
        try:
            input_path=Path(filedialog.askopenfile(mode='r',filetypes=[("PNG files","*.png")],
                                                   title="Select The PNG File.").name)
            output_path = Path(input_path).parent / (Path(input_path).stem + ".jpg")
            Image.open(input_path).save(output_path)
            print(termcolor.colored(text=f"\nFile Is Saved On '{output_path}' \n",color="green"))
        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
    
    def pdf_encrypt(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES","*.pdf")],
                                                     title="Select PDF File To Encrypt.").name)
            if self.check_is_encrypted(input_path):
                termcolor.cprint(text="\nSelected The Already Encrypted PDF File\n",color='red')
                return
            password1 = simpledialog.askstring(title="Choose Strong Password",prompt="Enter Password To Encrypt The File: ")
            if password1 == "" or password1 == "Cancle" or password1 == None:
                termcolor.cprint(text="\nPassword Must Require To Encrypt The File....\n",color='red')
                return
            password2=simpledialog.askstring(prompt="Enter Password Again: ",title="Rechecking password...",show='*')
            if password2 == "" or password2 == "Cancle" or password2 == None:
                termcolor.cprint(text="\nReverification Of Password Must Require To Encrypt The File\n",color='red')
                return
            if password1 != password2:
                print("\033[2J","\033[H",termcolor.colored(text=f"\nEntered Wrong Password For The Second Time '{password1}' != '{password2}' \n",color="red"))
                return
            
            pdf_writer = PyPDF2.PdfWriter()
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                output_folder = filedialog.askdirectory(title="Select Folder To Save The Encrypted PDF File.")
                if output_folder == ""  or output_folder == None:
                    termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                    return
                output_path = Path(output_folder) / (Path(input_path).stem + " encrypted.pdf")
                if " decrypted" in input_path.as_posix():
                    output_path = str(output_path).replace(" decrypted","")

                for i in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[i]
                    pdf_writer.add_page(page)
                pdf_writer.encrypt(user_password='0123581321',owner_password=password2)
                pdf_writer.write(output_path)
            termcolor.cprint(text=f"\nFile Is Saved on '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
    
    def add_watermark(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES","*.pdf")],
                                                     title="Select PDF File To ADD Water Mark.").name)
            if self.check_is_encrypted(input_path):
                termcolor.cprint(text="\nSelected The Encrypted File To ADD Water Mark File Must Be Decrypted\n",color='red')
                return
            water_mark_location = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES","*.pdf")],
                                                              title="Select Water Mark PDF File.").name)
            if self.check_is_encrypted(water_mark_location):
                termcolor.cprint(text="\nSelected The Encrypted File Decrypt The Water Mark File\n",color='red')
                return
            output_folder = filedialog.askdirectory(title="Select Folder To Save The Water Marked File.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            output_path = Path(output_folder) / (Path(input_path).stem + " watermarked.pdf")
            water_mark_file = open(water_mark_location,'rb')
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                pdf_writer = PyPDF2.PdfWriter()
                for i in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[i]
                    page.merge_page(PyPDF2.PdfReader(open(water_mark_file.name,'rb')).pages[0])
                    pdf_writer.add_page(page)
                pdf_writer.write(output_path)
            water_mark_file.close()
            termcolor.cprint(text=f"\nFile Is Saved on '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input/Water Mark File Is Selected\n",color='red')
        except Exception as e:
            print(e)
    
    def pdf_decrypt(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES",'*.pdf')],
                                                     title="Select Encrypted PDF File To Decrypt.").name)
            if not self.check_is_encrypted(input_path):
                termcolor.cprint(text="\nSelected The Normal PDF File Not Encrypted\n",color='red')
                return
            password = simpledialog.askstring(title="Enter Correct Password",prompt="Enter Password To Decrypt The File: ")
            if password == "" or password == "Cancle" or password == None:
                termcolor.cprint(text="\nPassword Must Require To Decrypt The File\n",color='red')
                return
            pdf_writer = PyPDF2.PdfWriter()
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                pdf_reader.decrypt(password=password)
                for i in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[i]
                    pdf_writer.add_page(page)
            output_folder = filedialog.askdirectory(title="Select Folder To Save The Decrypted PDF File.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color="red")
                return
            output_path = Path(output_folder) / (Path(input_path).stem + " decrypted.pdf")
            if " encrypted" in input_path.as_posix():
                output_path=str(output_path).replace(" encrypted","")
            pdf_writer.write(output_path)
        
            termcolor.cprint(text=f"\nFile Is Saved On The '{output_path}' \n",color="green")
        
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except PyPDF2.errors.FileNotDecryptedError:
            termcolor.cprint(text=f"\nYou Entered Wrong Password To Decrypt '{password}' \n",color='red')
        except Exception as e:
            print(e)
    
    def pdf_split(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES","*.pdf")],
                                                     title="Select The PDF File To Split.").name)
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                if pdf_reader.is_encrypted:
                    termcolor.cprint(text='\nSelected The Encrypted PDF Decrypt The PDF\n',color='red')
                    return
                max_pages = len(pdf_reader.pages)
            if max_pages < 2:
                termcolor.cprint(text="\nFor Spliting Min Pages Should Be 2\n",color='red')
                return
            output_folder = filedialog.askdirectory(title='Select The Folder To Save The Splited PDF File.')
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            split_page_no = simpledialog.askinteger(title="Getting How Many Split's.",prompt=f"Enter Page No Where You Want To Split PDF\nFor FULL PDF Enter '{max_pages+1}'\nFor Range Enter '0' ",minvalue=0,maxvalue=max_pages+1)
            if split_page_no == None:
                termcolor.cprint(text="\nEnter The Page No To Get PDF\n",color='red')
                return
            elif split_page_no == max_pages+1: #means get splilleted every pdf page
                with open(input_path,'rb') as fo:
                    pdf_reader = PyPDF2.PdfReader(fo)
                    for i in range(len(pdf_reader.pages)):
                        page=pdf_reader.pages[i]
                        pdf_writer=PyPDF2.PdfWriter()
                        pdf_writer.add_page(page)
                        pdf_writer.write(os.path.join(output_folder,(f"{i+1:02}.pdf"))) # upto 2 digits refer format strings
                termcolor.cprint(text=f"\nFiles Are Saved On '{output_folder}' Folder\n",color='green')
            elif split_page_no == 0: #split for range
                page_nos = simpledialog.askstring(title="Getting Page No",prompt=f"Enter Page No's spaced by ',' Where You Want The PDF")
                if page_nos == None:
                    termcolor.cprint(text="\nEnter The Page No To Get PDF\n",color='red')
                    return
                else:
                    try:
                        page_list = list(map(int,page_nos.split(',')))
                        assert all(1 <= page <= max_pages for page in page_list), f"\nPage numbers must be between 1 and {max_pages}\n"
                    except AssertionError as e:
                        print(e)
                        return
                    output_path = Path(output_folder) / (Path(input_path).stem + " splited pages.pdf")
                    with open(input_path,'rb') as fo:
                        pdf_reader = PyPDF2.PdfReader(fo)
                        pdf_writer = PyPDF2.PdfWriter()
                        for page in page_list:
                            pdf_writer.add_page(pdf_reader.pages[page-1])
                        pdf_writer.write(output_path)
                    termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' Folder\n",color='green')
                    
            else: #for particular page
                output_path = Path(output_folder) / (Path(input_path).stem + f" {split_page_no}.pdf")
                with open(input_path,'rb') as fo:
                    pdf_reader = PyPDF2.PdfReader(fo)
                    pdf_writer = PyPDF2.PdfWriter()
                    pdf_writer.add_page(pdf_reader.pages[split_page_no-1])
                    pdf_writer.write(output_path)
                termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
            

    def pdf_split_half(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES","*.pdf")],
                                                     title="Select The PDF File To Split.").name)
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                if pdf_reader.is_encrypted:
                    termcolor.cprint(text='\nSelected The Encrypted PDF To Split Decrypt The PDF\n',color='red')
                    return
                max_pages = len(pdf_reader.pages)
            if max_pages < 2:
                termcolor.cprint(text="\nFor Spliting Min Pages Should Be 2\n",color='red')
                return
            output_folder = filedialog.askdirectory(title='Select The Folder To Save The Splited PDF File.')
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            split_page_no = simpledialog.askinteger(title="Getting Page No",prompt="Enter Page No Where You\n Want To Split Into Two",minvalue=2,maxvalue=max_pages)
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                pdf_writer1 = PyPDF2.PdfWriter()
                pdf_writer2 = PyPDF2.PdfWriter()
                l = pdf_reader.pages
                part_1 = l[:split_page_no]
                part_2 = l[split_page_no:]
                for i in part_1:
                    page = pdf_reader.pages[i]
                    pdf_writer1.add_page(page)
                pdf_writer1.write(Path(output_folder) / (Path(input_path).stem) + " Part-1.pdf")
                for i in part_2:
                    page = pdf_reader.pages[i]
                    pdf_writer2.add_page(page)
                pdf_writer2.write(Path(output_folder) / (Path(input_path).stem) + " Part-2.pdf")
                termcolor.cprint(text=f"\nFiles Are Saved On '{output_folder}' Folder\n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
    
    def pdf_merger(self):
        try:
            input_path = Path(filedialog.askdirectory(title="Select Folder Contains All PDF's Only."))
            conforamtion = pymsgbox.confirm(text="Make The Pages In Sort Order To Get Expected Result",title='WARNING')
            if conforamtion == 'Cancel':
                termcolor.cprint(text="\nPlease Sort The Pages Order Like Page No 01,02....\n",color='red')
                return
            pdf_list = input_path.glob("*.*pdf")
            pdf_list = [Path(i).as_posix() for i in pdf_list]
            for i in pdf_list:
                if self.check_is_encrypted(i):
                    termcolor.cprint(text=f"\nFile '{i}' Is Encrypted Decrypt The File To Further\n",color='red')
                    return
            pdf_list = sorted(pdf_list)
            termcolor.cprint(text=pdf_list,color='green')
            output_folder = filedialog.askdirectory(title='Select Folder To Save The Merged PDF.')
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color="red")
                return
            file=Path(filedialog.asksaveasfile(mode='w',filetypes=[("PDF files","*.pdf")],defaultextension=".pdf",title="Enter The File Name To Save").name)
            if file == "" or file == None:
                termcolor.cprint(text="\nMust Enter The Name To Save The File\n",color='red')
                return
            pdf_merge = PyPDF2.PdfMerger()
            for j in pdf_list:
                pdf_merge.append(j)
            output_path = os.path.join(output_folder,file)
            pdf_merge.write(output_path)
            termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color="green")
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except PyPDF2.errors.EmptyFileError:
            print(termcolor.colored(f"\nSeletced The Empty File name '{os.path.join(output_folder,i)}' \n",color='red'))
        except Exception as e:
            print(e)

    def pdf_compress(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[('PDF FILES',"*.pdf")],
                                                     title="Select PDF File To Compress.").name)
            if self.check_is_encrypted(input_path):
                termcolor.cprint(text='\nSelected The Encrypted PDF To Compress Decrypt The PDF\n',color='red')
                return
            output_folder = filedialog.askdirectory(title='Select The Folder To Save The Compressed PDF File.')
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            key = 'project_public_bdfcd65ef641f193fc4aac801e8f34ac_ZHywn223d2e4761d1362c39016d7371e50c5d' #api key
            pdf = pdf_compressor.Task(public_key=key,tool='compress')
            pdf.add_file(file_path=input_path)
            pdf.upload()
            pdf.process()
            pdf.download(save_to_dir=output_folder)
            output_path = Path(output_folder) / (Path(input_path).stem + " compressed.pdf")
            termcolor.cprint(text=f"\nFile Is Saved on '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except requests.exceptions.ConnectionError:
            termcolor.cprint(text="\nNo Internet Connection........\n",color='red')
        except Exception as e:
            print(e)
    
    def addpageno_PDF(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF FILES",'*.pdf')],
                                                        title='Select PDF File To ADD Page No.').name)
            if self.check_is_encrypted(input_path):
                termcolor.cprint(text="\nSelected The Encrypted PDF To Add PageNo's Decrypt The PDF\n",color='red')
                return
            output_folder = filedialog.askdirectory(title="Select The Folder To Save The PDF With Pange No.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            
            key = 'project_public_bdfcd65ef641f193fc4aac801e8f34ac_ZHywn223d2e4761d1362c39016d7371e50c5d' #api key
            pdf = pdf_compressor.Task(public_key=key,tool='pagenumber')
            pdf.add_file(file_path=input_path)
            pdf.upload()
            pdf.process()
            output_path = Path(output_folder) / (Path(input_path).stem + " page no.pdf")
            pdf.download(output_path)
            termcolor.cprint(text=f"\nFile Is Saved on '{output_path}' \n",color='green')
        except requests.exceptions.ConnectionError:
            termcolor.cprint(text="\nNo Internet Connection........\n",color='red')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
    
    def get_text(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF Files","*.pdf")],
                                                     title="Select PDF File To Get TEXT From Page.").name)
            if self.check_is_encrypted(input_path):
                termcolor.cprint(text="\nSelected The Encrypted PDf TO Get TEXT Decrypt The PDF\n",color='red')
                return
            output_folder = filedialog.askdirectory(title="Select Folder To Save TEXT Of PDF.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                for i in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[i]
                    text = page.extract_text()
                    with open(Path(output_folder)/(Path(input_path).stem + f" page no {i+1}.txt"),'w') as f:
                        f.write(text)
    
            conformation = pymsgbox.confirm(text="Do You Want To Merger All TXT Files Into Single File",title="Conformation oF Merging TXT Files.")
            converted_list = [i.as_posix() for i in Path(output_folder).glob(f"{Path(input_path).stem} page no*.*txt")]
            if conformation == "OK":
                converting_files.merger_txt(input_path=output_folder,li=converted_list)
        
            termcolor.cprint(text=f"\nFiles Are Saves At {output_folder} \n",color='green')
        except UnicodeError:
            for i in converted_list:
                os.remove(i) # removes the converted txt files
            termcolor.cprint(text="\nUnable To Extract Text From Some Pages\n",color='red')
        except FileExistsError:
            path = Path(output_folder)/(Path(input_path).stem + f" page no {i}.txt")
            termcolor.cprint(text=f"\nFiles Already Exist's In Folder '{path}' \n",color='red')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)
    
    def merger_txt(self,input_path=None,li=None):
        try:
            output_path = Path(input_path) / Path("merged.pdf")
            if input_path == None: #Means For Single files 
                input_path = filedialog.askdirectory(title="Select TXT Files Folder To Merge TXT.")
                output_folder = filedialog.askdirectory(title="Select Folder To Save Merged TXT Files.")
                if output_folder == "" or output_folder == None:
                    termcolor.cprint(text="\nNo Output Fodler Is Selected....\n",color='red')
                file_name = Path(filedialog.asksaveasfile(mode='w',filetypes=[("TXT FILES",'*.txt')],
                                                     defaultextension=".txt",title="Enter The File Name To Save.").name)
                output_path = Path(input_path) / Path(file_name)
            if not li:
                txt_list = Path(input_path).glob("*.*txt")
            else:
                txt_list = li
            with open(output_path,'a') as fo:
                for i in txt_list:
                    with open(i,'r') as f:
                        text = f.read()
                    fo.write(text)
            termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except FileExistsError:
            termcolor.cprint(text=f"\nFilse Already Exist's In Folder '{output_path}' \n",color='red')
        except Exception as e:
            print(e)
            termcolor.cprint(text=f"\nGet An Exception We Removed File >'{output_path}' From System \n",color='red')
    
    def rotate_pdf(self):
        try:
            input_path = Path(filedialog.askopenfile(mode='r',filetypes=[("PDF Files","*.pdf")],
                                                     title="Select PDF File To Rotate.").name)
            with open(input_path,'rb') as fo:
                pdf_reader = PyPDF2.PdfReader(fo)
                if pdf_reader.is_encrypted:
                    termcolor.cprint(text="\nSelected The Encrypted PDF To Get Rotated Pages Decrypt The PDF\n",color='red')
                    return
                max_pages = len(pdf_reader.pages)
            page_no = simpledialog.askinteger(title="Getting Page NO.",prompt="Enter Page No Where To Rotate\nFor Full PDF Rotation Enter 0",minvalue=0,maxvalue=max_pages)
            if page_no == None:
                termcolor.cprint(text="\nEnter The Page No To Get PDF \n",color='red')
                return
            rotation_angle = simpledialog.askinteger(title="Getting Angle.",prompt="Enter Angle in degrees And Angle b/w 90 to 360 Must Be Divisible With 90",minvalue=90,maxvalue=360)
            if rotation_angle == None:
                termcolor.cprint(text="\nEnter The Page No To Get PDF \n",color='red')
                return
            if rotation_angle%90 != 0 :
                termcolor.cprint(text=f"\nEnter The Angle Divisible By 90 Not '{rotation_angle}' \n",color='red')
                return
            output_folder = filedialog.askdirectory(title=f"Select Folder To Save {page_no} Page Of PDF.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            pdf_writer = PyPDF2.PdfWriter()
            if page_no == 0: # means rotate all PDF
                output_path = Path(output_folder) / (Path(input_path).stem + " rotated.pdf")
                with open(input_path,'rb') as fo:
                    pdf_reader = PyPDF2.PdfReader(fo)
                    pages = pdf_reader.pages
                    for i in range(len(pages)):
                        page = pdf_reader.pages[i]
                        rotated_page = page.rotate(rotation_angle)
                        pdf_writer.add_page(rotated_page)
                    
                    pdf_writer.write(output_path)

            else:
                output_path = Path(output_folder) / (Path(input_path).stem + f" rotated {page_no}.pdf")
                with open(input_path,'rb') as fo:
                    pdf_reader = PyPDF2.PdfReader(fo)
                    pages = pdf_reader.pages
                    for i in range(len(pages)):
                        page = pages[i]
                        if i == page_no-1:
                            pdf_writer.add_page(page).rotate(rotation_angle)
                        else:
                            pdf_writer.add_page(page)
                pdf_writer.write(output_path)
            termcolor.cprint(text=f"\nFile Saved On '{output_path}' \n",color='green')
        except AttributeError:
            termcolor.cprint(text="\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)

    def multiple_files(self):
        x = simpledialog.askinteger(title="Select For Multipul Files",prompt="1.PDF Files\n2.Word Files\n",minvalue=1,maxvalue=2)
        match x:
            case 1:
                input_path = filedialog.askdirectory(title="Select PDF Files Folder.")
                if input_path is None or input_path == "":
                    termcolor.cprint(text="\nNo Input Folder Is Selected....\n",color='red')
                    return
                ouput_folder = filedialog.askdirectory(title="Select Folder To Save The Converted Word Files.")
                if ouput_folder is None or ouput_folder == "":
                    termcolor.cprint(text='\nNo Output Folder Is Selected....\n',color='red')
                    return
                self.pdf_docx(input_path=input_path,output_folder=ouput_folder)
            case 2:
                input_path = filedialog.askdirectory(title="Select DOCX Files Folder.")
                if input_path is None or input_path == "":
                    termcolor.cprint(text="\nNo Input Folder Is Selected....\n",color='red')
                    return
                ouput_folder = filedialog.askdirectory(title="Select Folder To Save The Converted PDF Files.")
                if ouput_folder is None or ouput_folder == "":
                    termcolor.cprint(text='\nNo Output Folder Is Selected....\n',color='red')
                    return
                self.docx_pdf(input_path=input_path,output_folder=ouput_folder)
            case _:
                termcolor.cprint(text="\nNot Valid Enter Only 1 or 2....\n",color='red')
                return
    def img_sketch(self):
        try:
            input_path = filedialog.askopenfile(mode='r',title="Select JPG/PNG File Only.",
                                                     filetypes=[("PNG FILES",'*.png'),("JPG Files","*.jpg")]).name
            output_folder = filedialog.askdirectory(title="Select The Folder To Save PDF.")
            if output_folder == "" or output_folder == None:
                termcolor.cprint(text="\nNo Output Folder Is Selected....\n",color='red')
                return
            output_path = Path(Path(output_folder) / (Path(input_path).stem + " sketch.png")).as_posix()

            image = cv2.imread(filename=input_path)
            gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            inverted_gray_image = 255 - gray_image
            blurred_image = cv2.GaussianBlur(inverted_gray_image, (21, 21), 0)
            inverted_blurred_image = 255 - blurred_image
            sketch_image = cv2.divide(gray_image, inverted_blurred_image, scale=256.0)
            cv2.imwrite(filename=output_path,img=sketch_image)
            
            termcolor.cprint(text=f"\nFile Is Saved On '{output_path}' \n",color='green')

        except AttributeError:
            termcolor.cprint(text=f"\nNo Input File Is Selected....\n",color='red')
        except Exception as e:
            print(e)

    def refresh(self):
        try:  # if only internet connection is avilable the buttons will activate
            if requests.get("https://developer.ilovepdf.com/user#/all/all/task/last24h").status_code == 200:
                ALLBUTTONS(x=200,y=200,text="ADD PAGE NO'S TO PDF",bcolor="#2cfc03",fcolor="#141414",command=converting_files.addpageno_PDF)
                ALLBUTTONS(x=400,y=200,text="PDF COMPRESS",bcolor="#03cafc",fcolor="#141414",command=converting_files.pdf_compress)
        except:
            ALLBUTTONS(x=200,y=200,text="ADD PAGE NO'S TO PDF",bcolor="#141414",fcolor="#141414",command=converting_files.addpageno_PDF,state='disabled')
            ALLBUTTONS(x=400,y=200,text="PDF COMPRESS",bcolor="#141414",fcolor="#141414",command=converting_files.pdf_compress,state='disabled')
    def recursive(self,path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath('.')
        return os.path.join(base_path, path)



#### MAIN BUTTONS ####
converting_files = Converting_files()
def ALLBUTTONS(x,y,text,bcolor,fcolor,command,state='active'):
    def on_entry(e):
        mybutton['background'] = bcolor
        mybutton['foreground'] = fcolor
    def on_leave(e):
        mybutton['background'] = fcolor
        mybutton['foreground'] = bcolor

    mybutton = Button(root,width=25,height=2,text=text
                      ,fg=bcolor,bg=fcolor,border=5,
                      activebackground=bcolor,activeforeground=fcolor,
                      command=command,state=state)
    mybutton.bind('<Enter>',on_entry)
    mybutton.bind('<Leave>',on_leave)
    mybutton.place(x=x,y=y)


try:  # if only internet connection is avilable the buttons will activate
    if requests.get("https://developer.ilovepdf.com/user#/all/all/task/last24h").status_code == 200:
        ALLBUTTONS(x=200,y=200,text="ADD PAGE NO'S TO PDF",bcolor="#2cfc03",fcolor="#141414",command=converting_files.addpageno_PDF)
        ALLBUTTONS(x=400,y=200,text="PDF COMPRESS",bcolor="#03cafc",fcolor="#141414",command=converting_files.pdf_compress)
except:
    ALLBUTTONS(x=200,y=200,text="ADD PAGE NO'S TO PDF",bcolor="#141414",fcolor="#141414",command=converting_files.addpageno_PDF,state='disabled')
    ALLBUTTONS(x=400,y=200,text="PDF COMPRESS",bcolor="#141414",fcolor="#141414",command=converting_files.pdf_compress,state='disabled')
finally:
    root.iconbitmap(bitmap=converting_files.recursive(r"C:\icons\converting.ico"))
    ALLBUTTONS(x=0,y=0,text="PDF 2 DOCX",bcolor="#a83e32",fcolor="#141414",command=converting_files.pdf_docx)
    ALLBUTTONS(x=200,y=0,text="DOCX 2 PDF",bcolor="#a86932",fcolor="#141414",command=converting_files.docx_pdf)
    ALLBUTTONS(x=400,y=0,text="IMG 2 PDF",bcolor="#a2a832",fcolor="#141414",command=converting_files.img_pdf)
    ALLBUTTONS(x=0,y=50,text="JPEG 2 PNG",bcolor="#5fa832",fcolor="#141414",command=converting_files.jpeg_png)
    ALLBUTTONS(x=200,y=50,text="PNG 2 JPEG",bcolor="#32a863",fcolor="#141414",command=converting_files.png_jpeg)
    ALLBUTTONS(x=400,y=50,text="PDF ENCRYPT",bcolor="#32a890",fcolor="#141414",command=converting_files.pdf_encrypt)
    ALLBUTTONS(x=0,y=100,text="PDF DECRYPT",bcolor="#3281a8",fcolor="#141414",command=converting_files.pdf_decrypt)
    ALLBUTTONS(x=200,y=100,text="ADD WATER MARK",bcolor="#324aa8",fcolor="#141414",command=converting_files.add_watermark)
    ALLBUTTONS(x=400,y=100,text="PDF SPLIT",bcolor="#5332a8",fcolor="#141414",command=converting_files.pdf_split)
    ALLBUTTONS(x=0,y=150,text="PDF MERGE",bcolor="#9232a8",fcolor="#141414",command=converting_files.pdf_merger)
    ALLBUTTONS(x=200,y=150,text="EXTRACT TEXT FROM PDF",bcolor="#8132a8",fcolor="#141414",command=converting_files.get_text)
    ALLBUTTONS(x=400,y=150,text="MERGE TEXT FILES",bcolor="#a0a832",fcolor="#141414",command=converting_files.merger_txt)
    ALLBUTTONS(x=0,y=200,text="ROTATE PDF BY ANGLE",bcolor="#a83232",fcolor="#141414",command=converting_files.rotate_pdf)
    ALLBUTTONS(x=0,y=250,text="MULTIPUL FILES",bcolor="#a83275",fcolor="#141414",command=converting_files.multiple_files)
    ALLBUTTONS(x=200,y=250,text="IMG 2 SKETCH",bcolor="#a83275",fcolor="#141414",command=converting_files.img_sketch)
    ALLBUTTONS(x=400,y=250,text="Refresh",bcolor="#a83275",fcolor="#141414",command=converting_files.refresh)
root.mainloop()

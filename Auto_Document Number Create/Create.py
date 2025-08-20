import os
import json
import win32com.client as win32
from datetime import datetime
import time

class DocumentNumberGenerator:
    def __init__(self):
        self.data_file = "contract_numbers.json"
        self.word_app = None
        self.template_path = r"rashid.baghirov\Desktop\Payriff_Documents_Generator\Payriff_Bakubus_Validator.docx"  #Changable
        self.output_folder = r"rashid.baghirov\Desktop\Payriff_Documents_Generator"                                 #Changable
        self.load_data()
        self.init_word()
        self.show_available_printers()
    
    def init_word(self):
        """Start Word application"""
        try:
            self.word_app = win32.Dispatch("Word.Application")
            self.word_app.Visible = False  
            print(" Word application started successfully")
        except Exception as e:
            print(f" Word application failed to start: {e}")
    
    def load_data(self):
        """Load existing contract numbers"""
        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            print(" Data loaded")
        except FileNotFoundError:
            self.data = {}
            print(" New data file created")
    
    def save_data(self):
        """Save data"""
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def generate_sequential_contract_number(self, start_num=120):
        """Sequential number generation starting from 120"""
        today = datetime.now().strftime("%Y%m%d")
        
        sequential_key = f"sequential_{today}"
        
        if sequential_key not in self.data:
            self.data[sequential_key] = {"current_number": start_num}
        
        current_num = self.data[sequential_key]["current_number"]
        contract_number = f"MQ-{today}-{current_num:03d}"
        
     
        self.data[sequential_key]["current_number"] = current_num + 1
        self.save_data()
        
        return contract_number, current_num
    
    def show_available_printers(self):
        """Show available printers"""
        print("\n" + "="*60)
        print(" PRINTER CHECK - DETAILED SEARCH")
        print("="*60)
        
       
        all_printers = self.get_all_printers_advanced()
        physical_printers = self.get_physical_printers()
        
        print(f"\n All found printers ({len(all_printers)} count):")
        for i, printer in enumerate(all_printers, 1):
            print(f"   {i}. {printer}")
        
        print(f"\n Physical printers ({len(physical_printers)} count):")
        if physical_printers:
            for i, printer in enumerate(physical_printers, 1):
                print(f"    {i}. {printer}")
            print(f"\n Selected printer: {physical_printers[0]}")
        else:
            print("     No physical printer found!")
            if all_printers:
                print(f"\n First available printer will be used: {all_printers[0]}")
        
        print("="*60)
    
    def get_all_printers_advanced(self):
        """Find all printers"""
        printers = []
        
        try:
          
            import subprocess
            print("🔍 Searching printers with WMIC...")
            result = subprocess.run(['wmic', 'printer', 'get', 'name,portname,drivername'], 
                                  capture_output=True, text=True, timeout=10)
            
            lines = result.stdout.split('\n')
            for line in lines[1:]:  
                line = line.strip()
                if line and line != 'Name':
                    parts = line.split()
                    if parts:
                        printer_name = ' '.join(parts[:-2]) if len(parts) > 2 else line
                        if printer_name and printer_name != 'Name':
                            printers.append(printer_name)
                            print(f"    Found: {printer_name}")
        except Exception as e:
            print(f" WMIC error: {e}")
        
        try:
         
            print(" Searching printers with PowerShell...")
            ps_command = "Get-Printer | Select-Object Name"
            result = subprocess.run(['powershell', '-Command', ps_command], 
                                  capture_output=True, text=True, timeout=10)
            
            lines = result.stdout.split('\n')
            for line in lines[2:]:
                line = line.strip()
                if line and line != '----' and line != 'Name':
                    printers.append(line)
                    print(f"   PowerShell found: {line}")
        except Exception as e:
            print(f" PowerShell error: {e}")
        
        try:
           
            print(" Searching printers via Word...")
            if self.word_app:
               
                doc = self.word_app.Documents.Add()
                
               
                dialogs = self.word_app.Dialogs
                print_dialog = dialogs.Item(88)  
                
                doc.Close()
                print("    Word printer dialog access available")
        except Exception as e:
            print(f" Word printer error: {e}")
        
        return list(set(printers))
    
    def get_physical_printers(self):
        """Find physical printers (exclude PDF, FAX, XPS)"""
        try:
            all_printers = self.get_all_printers_advanced()
            
          
            excluded_keywords = ['PDF', 'FAX', 'XPS', 'MICROSOFT', 'ONENOTE', 'SEND TO', 'VIRTUAL']
            
            physical_printers = []
            for printer in all_printers:
                is_physical = True
                printer_upper = printer.upper()
                
                for keyword in excluded_keywords:
                    if keyword in printer_upper:
                        is_physical = False
                        break
                
                if is_physical and printer.strip():
                    physical_printers.append(printer.strip())
            
            return physical_printers
        except Exception as e:
            print(f" Printer search error: {e}")
            return []
    
    def print_document(self, filepath):
        """Send document to real printer (without opening PDF dialog)"""
        try:
            if not os.path.exists(filepath):
                print(f" File to print not found: {filepath}")
                return False
                
            print(f" Print preparing: {os.path.basename(filepath)}")
            
         
            print(" Opening file for print...")
            doc = self.word_app.Documents.Open(filepath)
            
         
            physical_printers = self.get_physical_printers()
            all_printers = self.get_all_printers_advanced()
            
            print(" Sending print command...")
            time.sleep(1)
            
            success = False
            
         
            if physical_printers:
                for printer in physical_printers:
                    try:
                        print(f" Attempting {printer}...")
                        
                       
                        original_printer = self.word_app.ActivePrinter
                        self.word_app.ActivePrinter = printer
                        
                        
                        doc.PrintOut(
                            Background=False,
                            PrintToFile=False,
                            Copies=1,
                            Range=0  
                        )
                        
                        print(f" Sent to {printer}")
                        success = True
                        break
                        
                    except Exception as e:
                        print(f" {printer} error: {e}")
                        continue
            
         
            if not success and all_printers:
                for printer in all_printers:
                    if 'PDF' not in printer.upper():
                        try:
                            print(f" Attempting {printer}...")
                            
                            self.word_app.ActivePrinter = printer
                            doc.PrintOut(
                                Background=False,
                                PrintToFile=False,
                                Copies=1,
                                Range=0
                            )
                            
                            print(f" Sent to {printer}")
                            success = True
                            break
                            
                        except Exception as e:
                            print(f" {printer} error: {e}")
                            continue
            
         
            if not success:
                print(" Automatic print failed - manual attempt...")
                try:
                
                    doc.PrintOut(
                        Background=False,
                        PrintToFile=False,
                        Copies=1
                    )
                    print(" Manual print sent")
                    success = True
                except Exception as e:
                    print(f" Manual print error: {e}")
            
            print(" Waiting for print command to process...")
            time.sleep(3)
            
          
            doc.Close()
            print(" Print document closed")
            
            if success:
                print(f" Print command completed successfully!")
                print(" Check print queue")
            else:
                print(" Print failed")
            
            return success
            
        except Exception as e:
            print(f" Print error: {e}")
            try:
                if 'doc' in locals():
                    doc.Close()
                    print(" Print document forcibly closed")
            except:
                pass
            return False
    
    def update_word_document(self, contract_number, auto_print=True):
        """Update document_number bookmark in Word document and save with new name"""
        try:
            if not os.path.exists(self.template_path):
                print(f" Error: Template file not found: {self.template_path}")
                return False, None
            
            print(f" Opening Word document...")
            
           
            doc = self.word_app.Documents.Open(self.template_path)
            print(f" Template opened: {os.path.basename(self.template_path)}")
            
          
            if doc.Bookmarks.Exists("document_number"):
                bookmark = doc.Bookmarks("document_number")
                bookmark.Range.Text = str(contract_number)
                print(f" Bookmark 'document_number' updated: {contract_number}")
            else:
                print(" Error: 'document_number' bookmark not found!")
                doc.Close()
                return False, None
            
          
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_filename = f"Contract_{contract_number}_{timestamp}.docx"
            new_filepath = os.path.join(self.output_folder, new_filename)
            
            print(f" Saving document: {new_filename}")
         
            doc.SaveAs2(new_filepath)
            print(f" File saved: {new_filepath}")
            
            
            doc.Close()
            print(f" Word document closed")
            
           
            if os.path.exists(new_filepath):
                file_size = os.path.getsize(new_filepath)
                print(f" File confirmed: {file_size} bytes")
                
                
                if auto_print:
                    try:
                        print_success = self.print_document(new_filepath)
                        return True, new_filepath
                    except Exception as print_error:
                        print(f" Print error: {print_error}")
                        return True, new_filepath  
                else:
                    return True, new_filepath
            else:
                print(" Error: File was not saved!")
                return False, None
            
        except Exception as e:
            print(f" Error while updating Word document: {e}")
            try:
                if 'doc' in locals():
                    doc.Close()
                    print(" Word document forcibly closed")
            except:
                pass
            return False, None
    
    def auto_generate_and_save(self):
        """Automatically generates sequential number and saves Word document and prints"""
        print("\n" + "="*50)
        
        result = self.generate_sequential_contract_number(120)
        
        if result[0] is None:
            print(" Contract number could not be generated!")
            return False
        
        contract_number, number_only = result
        print(f" Contract number: {contract_number}")
        print(f" Number: {number_only}")
        
       
        success, filepath = self.update_word_document(number_only, auto_print=True)
        
        if success and filepath:
            print(f" Contract created successfully!")
            print(f" File path: {filepath}")
            return True
        else:
            print(" Contract could not be created!")
            return False
    
    def close_word(self):
        """Close Word application"""
        try:
            if self.word_app:
                self.word_app.Quit()
                print(" Word application closed")
        except:
            pass
    
    def run_auto_mode(self):
        """Automatically runs starting from 120 until 40 files are created"""
        try:
            print(" CREATING AND PRINTING 1 CONTRACT")
            print(" Starting from 120 with sequential increment")
            print(" Each file is automatically printed")
            print("="*60)
            
            target_count = 1
            success_count = 0
            failed_count = 0
            
            for i in range(1, target_count + 1):
                print(f"\n {i}/{target_count} - Creating contract...")
                
                success = self.auto_generate_and_save()
                
                if success:
                    success_count += 1
                    print(f" {i}/{target_count} successful!")
                else:
                    failed_count += 1
                    print(f" {i}/{target_count} failed!")
                
             
                if i < target_count:
                    print(" Waiting 2 seconds...")
                    time.sleep(2)
            
            
            print("\n" + "="*60)
            print(" FINAL RESULTS")
            print("="*60)
            print(f" Target: {target_count} files")
            print(f" Successful: {success_count} files")
            print(f" Failed: {failed_count} files")
            print("="*60)
                
        except KeyboardInterrupt:
            print(f"\n⏹ Stopped by user")
            print(f" {success_count} files created successfully")
        except Exception as e:
            print(f"\n Unexpected error: {e}")
        finally:
            self.close_word()
            print(" Program completed")

if __name__ == "__main__":
    generator = DocumentNumberGenerator()
    
    # Run automatic mode
    generator.run_auto_mode()
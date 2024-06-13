import re
import time
import datetime
import os
import pythoncom
from pyautocad import Autocad
from pybricscad import Bricscad
from pyautocad.cache import Cached
import customtkinter as ctk

class GDPutils:

    def __init__(self, log_textbox=None):
        self.log_textbox = log_textbox
        self.log_file_path = self.create_log_file_with_initial_log()
        

    def create_log_file_with_initial_log(self):
        try:
            # Create the logs directory if it doesn't exist
            logs_dir = 'logs'
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)

            # Get the current date and time
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

            # Construct the filename
            filename = f"log_{timestamp}.txt"
            filepath = os.path.join(logs_dir, filename)

            # Write some initial log messages
            initial_log_messages = [
                "=== Log file created ===",
                f"Date and time: {now}",
                "Initial log messages:"
                "..."
            ]

            # Write the initial log messages to the file
            with open(filepath, 'w') as file:
                file.write('\n'.join(initial_log_messages) + '\n')

            print(f"Log file '{filename}' created with initial log messages.")
            return filepath
        except Exception as e:
            print(f"Error creating log file: {e}")    

    def log_to_box(self, message):
        """Log messages to the log textbox."""
        if self.log_textbox != None:
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert(ctk.END, message + '\n')
            self.log_textbox.see(ctk.END)  # Auto-scroll to the end
            self.log_textbox.configure(state="disabled")

    def log_to_file(self, message):
        """Log text to the log file."""
        try:
            with open(self.log_file_path, 'a') as file:
                file.write(message + '\n')
        except Exception as e:
            print(f"Error logging to file: {e}")

    def log(self, message):
        """Log messages to the log textbox and log file."""
        try:
            self.log_to_box(message)
            self.log_to_file(message)
        except Exception as e:
            print(f"Error logging: {e}")
    



    def replace_cad_text(self, find_replace_dict, CAD):
        if CAD == "AutoCAD":
            self.replace_acad_text(find_replace_dict)
        elif CAD == "BricsCAD":
            self.replace_bcad_text(find_replace_dict)
        else:
            self.log("Versão não implementada")

    def replace_acad_text(self, find_replace_dict):
        """Replace text in AutoCAD drawing."""
        try:
            self.log("Iniciando interface com o AutoCAD...")
            pythoncom.CoInitialize()
            acad = Autocad()
            if acad.doc is None:
                raise Exception("Falha ao abrir interface com o AutoCAD, verifique se o AutoCAD está aberto.")
            self.log("Interface com o AutoCAD iniciada com sucesso.")
            self.log("Abrindo desenho: " + acad.doc.Name)
            self.log("Varrendo códigos...")
            start_time = time.time()
            replaced_count = 0

            # Construct a single regular expression pattern
            pattern = re.compile("|".join(re.escape(k) for k in find_replace_dict.keys()))

            acad_cached = Cached(acad)
            updates = []  # To store updates to be made
            log_entries = []  # To store log entries to be made

            for text in acad_cached.iter_objects_fast('Text'):
                # Use the pattern to find all occurrences of any pattern to replace
                new_text = pattern.sub(lambda x: find_replace_dict[x.group(0)], text.TextString)

                if new_text != text.TextString:
                    updates.append((text, new_text))
                    log_entries.append(f"\nSubstituído\n{text.TextString}\npara \n{new_text}")
                    replaced_count += 1
            elapsed_time = time.time() - start_time
            self.log(f"Varredura concluída em {elapsed_time:.2f} segundos. Achado {replaced_count} para substituição.")
            # Apply all updates
            self.log("Aplicando atualizações...")
            for text, new_text in updates:
                text.TextString = new_text

            # Log all changes
            for log_entry in log_entries:
                self.log_to_file(log_entry)

            elapsed_time = time.time() - start_time
            self.log(f"Concluído em {elapsed_time:.2f} segundos. Substituído {replaced_count} ocorrências.")
        except Exception as e:
            print(f"Erro ao substituir texto: {e}")
            self.log(f"Erro ao substituir texto: {e}")


    def replace_acad_text_express_tools(self, find_replace_dict):
        """Replace text in AutoCAD drawing using AutoLISP script."""
        try:
            self.log("Iniciando interface com o AutoCAD...")
            pythoncom.CoInitialize()
            acad = Autocad()
            if acad.doc.Name is None:
                raise Exception("Falha ao abrir interface com o AutoCAD, verifique se o AutoCAD está aberto.")
            self.log("Interface com o AutoCAD iniciada com sucesso.")
            self.log("Abrindo desenho: " + acad.doc.Name)
            self.log("Substituindo códigos com AutoLISP...")

            # Define the AutoLISP function
            lisp_function = """
            (defun c:replaceText (txt2find txt2replace / ssText cnt e eg isMtext txt)
            (setq ssText (ssget "_X" '((0 . "*TEXT"))))
            (setq isMtext nil)
            (repeat (setq cnt (sslength ssText))
                (setq e (ssname ssText (setq cnt (1- cnt))))
                (setq eg (entget e))
                (if (eq (cdr (assoc 0 eg)) "MTEXT") (setq isMtext t) (setq isMtext nil))
                (if isMtext
                (setq txt (cdr (assoc 1 eg))) ; Get MTEXT content
                (setq txt (cdr (assoc 1 eg))) ; Get TEXT content
                ) ;if
                (setq txt (vl-string-subst txt2replace txt2find txt))
                (if isMtext
                (entmod (subst (cons 1 txt) (assoc 1 eg) eg)) ; Set MTEXT content
                (entmod (subst (cons 1 txt) (assoc 1 eg) eg)) ; Set TEXT content
                ) ;if
                (entupd e) ; Update the entity in the database
            ) ;repeat
            (prompt "\\nText Replaced!")
            )
            """
            
            # Load and execute the AutoLISP function
            acad.doc.SendCommand(lisp_function + "\n")

            start_time = time.time()
            replaced_count = 0

            # Call the AutoLISP function for each find and replace pair
            for find_text, replace_text in find_replace_dict.items():
                self.log(f"Substituindo '{find_text}' por '{replace_text}'")
                
                command_string = f'(replaceText "{find_text}" "{replace_text}")\n'
                acad.doc.SendCommand(command_string)

                replaced_count += 1

            elapsed_time = time.time() - start_time
            self.log(f"Concluído em {elapsed_time:.2f} segundos. Substituído {replaced_count} ocorrências.")
        except Exception as e:
            print(f"Erro ao substituir texto: {e}")
            self.log(f"Erro ao substituir texto: {e}")

    def replace_bcad_text(self, find_replace_dict):
        """Replace text in BricsCAD drawing."""
        try:
            self.log("Iniciando interface com o BricsCAD...")
            pythoncom.CoInitialize()
            bcad = Bricscad()
            if bcad.doc is None:
                raise Exception("Falha ao abrir interface com o BricsCAD, verifique se o BricsCAD está aberto.")
            self.log("Interface com o BricsCAD iniciada com sucesso.")
            self.log("Abrindo desenho: " + bcad.doc.Name)
            self.log("Varrendo códigos...")
            start_time = time.time()
            replaced_count = 0

            # Construct a single regular expression pattern
            pattern = re.compile("|".join(re.escape(k) for k in find_replace_dict.keys()))

            updates = []  # To store updates to be made
            log_entries = []  # To store log entries to be made

            for text in bcad.iter_objects_fast('Text'):
                # Use the pattern to find all occurrences of any pattern to replace
                new_text = pattern.sub(lambda x: find_replace_dict[x.group(0)], text.TextString)

                if new_text != text.TextString:
                    updates.append((text, new_text))
                    log_entries.append(f"\nSubstituído\n{text.TextString}\npara \n{new_text}")
                    replaced_count += 1

            elapsed_time = time.time() - start_time
            self.log(f"Varredura concluída em {elapsed_time:.2f} segundos. Achado {replaced_count} para substituição.")
            
            # Apply all updates
            self.log("Aplicando atualizações...")
            for text, new_text in updates:
                text.TextString = new_text
            

            
            # Refresh graphics in BricsCAD
            bcad.app.Update()  # Example update command, adjust as per BricsCAD API

            # Log all changes
            for log_entry in log_entries:
                self.log_to_file(log_entry)

            elapsed_time = time.time() - start_time
            self.log(f"Concluído em {elapsed_time:.2f} segundos. Substituído {replaced_count} ocorrências.")

        except Exception as e:
            self.log(f"Erro ao substituir texto: {e}")

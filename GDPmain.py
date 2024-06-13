import customtkinter as ctk
import threading, math, os, winsound
from datetime import datetime
from PIL import Image
from GDPutils import GDPutils

class ReplaceValuesApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("GDP")
        self.geometry("1090x800")
        
        # Define a 2D list of entries with their labels, variable names, and grid positions
        self.entries_info = [
            [
                 {"label": "Nome da SE:", "var_name": "se_name"},
                 {"label": "Sigla da SE:", "var_name": "se_code"},
                 {"label": "empty", "var_name": "empty"},
                 {"label": "Tensão da SE", "var_name": "se_voltage"}
            ],
            [
                {"label": "Cod. Equip.:", "var_name": "equip_code"},
                {"label": "Cod. TC:", "var_name": "tc_code"},
                {"label": "Cod. TP:", "var_name": "tp_code"},
                {"label": "empty", "var_name": "empty"}
            ],
            [
                {"label": "UCP:", "var_name": "ucp_value"},
                {"label": "PP:", "var_name": "pp_value"},
                {"label": "Numero do DJ do PP:", "var_name": "djpp_value"},
                {"label": "empty", "var_name": "empty"}
            ],
            [
                {"label": "UM:", "var_name": "um_value"},
                {"label": "PM:", "var_name": "pm_value"},
                {"label": "Numero do Medidor:", "var_name": "mpm_value"}               
            ],
            [
                {"label": "Chave 4:", "var_name": "ch4_code"},
                {"label": "Chave 5:", "var_name": "ch5_code"},
                {"label": "Chave 6:", "var_name": "ch6_code"},
                {"label": "Chave TP:", "var_name": "chTP_code"}
            ],
            [
                {"label": "DJ Al. CC:", "var_name": "DJCC_number"},
                {"label": "Borne CC Pos", "var_name": "BPCC_number"},
                {"label": "Borne CC Neg", "var_name": "BNCC_number"},
                {"label": "Cod. Aterramento", "var_name": "gnd_code"}
            ],
            [
                {"label": "DJ Al. CA:", "var_name": "DJCA_number"},
                {"label": "Borne CA Pos", "var_name": "BPCA_number"},
                {"label": "Borne CA Neg", "var_name": "BNCA_number"},
                {"label": "Cod. Al", "var_name": "al_code"}
            ],
            [
                {"label": "Primeiro nome do projetista:", "var_name": "pj_fname"},
                {"label": "Nome do projetista:", "var_name": "pj_name"},
                {"label": "Data:", "var_name": "date"}               
            ]
        ]

        self.voltages = ["138KV", "69KV", "34,5KV", "23KV", "13,8KV"]

        num_rows = 2*len(self.entries_info)
        num_columns = max(len(row) for row in self.entries_info)

        self.variables = {}
        self.entries = {}
        

        # Create and adjust the labels and images for the program
        logo_scale = 0.2
        logo_image = ctk.CTkImage(dark_image=Image.open("bin/SVpowerLogo.png"), size = (int(1245*logo_scale), int(230*logo_scale)))
        logo_label = ctk.CTkLabel(self, image=logo_image, text="")
        logo_label.place(relx=0.01, rely=0.015, anchor='nw')
        icon_image = ctk.CTkImage(dark_image=Image.open("bin/icon.png"), size = (35, 35))
        main_label = ctk.CTkLabel(self, text="Gerador de Projetos", image=icon_image, compound='left', font=("Roboto", 24))
        main_label.pack(pady=20, padx=10)
        
        version_label =  ctk.CTkLabel(self,text="V0.01", font=("Roboto", 10))
        version_label.place(relx=0.99, rely=0.999, anchor='se')

        creator_name =  ctk.CTkLabel(self,text="Feito por Lucas Spínola Bertão", font=("Roboto", 10))
        creator_name.place(relx=0.01, rely=0.999, anchor='sw')

        # Left Frame for Inputs
        left_frame = ctk.CTkFrame(self)
        left_frame.place(relx=0.01, rely=0.52, anchor='w')
        left_frame.grid_columnconfigure(num_columns, weight=1)
        left_frame.grid_rowconfigure(num_rows, weight=1)

        # Right Frame for log
        right_frame = ctk.CTkFrame(self)
        right_frame.place(relx=0.99, rely=0.52, anchor='e')

        self.log_textbox = ctk.CTkTextbox(right_frame, width=400, height=405, font=("Roboto", 20))
        self.log_textbox.pack(pady=12)
        self.log_textbox.configure(state="disabled")

        #Start GDPutils passing the log_box
        self.GDPutils = GDPutils(self.log_textbox)

       

        # Create widgets dynamically based on entries_info
        for row_idx, row_entries in enumerate(self.entries_info):
            for col_idx, entry_info in enumerate(row_entries):
                
                var_name = entry_info["var_name"]
                label_text = entry_info["label"]

                if label_text != "empty":

                    # Create StringVar and store it in a dictionary
                    self.variables[var_name] = ctk.StringVar()

                    # Create and place the label
                    label = ctk.CTkLabel(left_frame, text=label_text, anchor="sw", justify= "left")
                    label.grid(row=row_idx*2, column=col_idx, pady=5, padx=10)

                    # Create and place the entry
                    entry = ctk.CTkEntry(left_frame, textvariable=self.variables[var_name])
                    entry.grid(row=row_idx*2 + 1, column=col_idx, pady=3, padx=10)
                    
                    # Store the entry widget instance
                    self.entries[var_name] = entry

        self.replace_button = ctk.CTkButton(left_frame, text="Substituir Valores", command=self.replace_values_thread)
        self.replace_button.grid(row=len(self.entries_info)*2, columnspan=3, pady=20)
        self.load_default_info()


        #Entries customizadas:
        #Seletor de Estado
        self.estado_value = ctk.StringVar()
        estado_label = ctk.CTkLabel(left_frame, text="Estado:", anchor="sw", justify= "left")
        estado_label.grid(row=0, column = 2, pady=5, padx=10)

        self.estado_entry = ctk.CTkOptionMenu(left_frame, variable=self.estado_value, anchor="center", values=["Alagoas", "Amapá", "Maranhão", "Pará", "Piauí", "Rio Grande do Sul"], command=self.update_tc_code)
        self.estado_entry.set("Alagoas")
        self.estado_entry.grid(row=1, column = 2,pady=3, padx=10)

        #Seletor de tensão do equip
        self.voltage_value = ctk.StringVar()
        voltage_label = ctk.CTkLabel(left_frame, text="Tensão do equipamento:", anchor="sw", justify= "left")
        voltage_label.grid(row=6, column = 3, pady=5, padx=10)

        self.voltage_entry = ctk.CTkOptionMenu(left_frame, variable=self.voltage_value, anchor="center", values=self.voltages, command=self.update_tp_code)
        self.voltage_entry.set("13,8KV")
        self.voltage_entry.grid(row=7, column = 3,pady=3, padx=10)

        #Seletor de tensão da SE
        self.entries["se_voltage"].configure(state="disabled")
        voltageSel_frame = ctk.CTkFrame(left_frame)
        voltageSel_frame.configure(height = 75)
        voltageSel_frame.grid(row=2, column = 3, rowspan=4)
        voltageSel_frame.grid_columnconfigure(len(self.voltages), weight=1)
        voltageSel_frame.grid_rowconfigure(3, weight=1)
        checkboxsize = 1\


        self.checkbox_vars = {}
        for idx, voltage in enumerate(self.voltages):
            checkbox_var = ctk.BooleanVar()
            checkbox = ctk.CTkCheckBox(voltageSel_frame, text=voltage, variable=checkbox_var, height=checkboxsize, width=checkboxsize, command=self.update_voltage_entry)
            checkbox.grid(row=idx, column=0, sticky="w")
            self.checkbox_vars[voltage] = checkbox_var

        #Set commands on entry changes
        self.entries["equip_code"].bind("<KeyRelease>", lambda event: self.update_tc_code())
        self.entries["equip_code"].bind("<KeyRelease>", lambda event: self.update_ch_code())
        self.entries["tp_code"].bind("<KeyRelease>", lambda event: self.update_chTP_code())
        self.entries["ucp_value"].bind("<KeyRelease>", lambda event: self.on_ucp_change())
        self.entries["um_value"].bind("<KeyRelease>", lambda event: self.on_um_change())        
        self.entries["pj_fname"].bind("<KeyRelease>", lambda event: self.save_pj_info())          
        self.entries["pj_name"].bind("<KeyRelease>", lambda event: self.save_pj_info())


    def save_pj_info(self):
        """Save the values of pj_fname and pj_name to a file."""
        pj_fname = self.variables["pj_fname"].get()
        pj_name = self.variables["pj_name"].get()

        with open("pj_info.txt", "w") as file:
            file.write(f"{pj_fname}\n{pj_name}")

    def load_default_info(self):
        self.variables["tp_code"].set("81B1")
        self.variables["chTP_code"].set("41B1")
        """Load the values of pj_fname and pj_name from a file and set the current date."""
        if os.path.exists("pj_info.txt"):
            with open("pj_info.txt", "r") as file:
                lines = file.readlines()
                if len(lines) >= 2:
                    self.variables["pj_fname"].set(lines[0].strip())
                    self.variables["pj_name"].set(lines[1].strip())

        # Set the current date
        current_date = datetime.now().strftime("%d/%m/%y")
        self.variables["date"].set(current_date)

        
    def on_ucp_change(self):
        try:
            UCP = int(self.variables['ucp_value'].get())
            PP = math.floor((UCP-1)/8) + 1
            Pos_DJ = (UCP-1)%8 + 1
            self.variables['pp_value'].set(PP)
            self.variables['djpp_value'].set(Pos_DJ)
        except ValueError as e:
            self.variables['pp_value'].set(0)
            self.variables['djpp_value'].set(0)

    def on_um_change(self):
        try:
            UM = int(self.variables['um_value'].get())
            PM = math.floor((UM-1)/8) + 1
            Pos_Med = (UM-1)%8 + 1
            self.variables['pm_value'].set(PM)
            self.variables['mpm_value'].set(Pos_Med)
        except ValueError as e:
            self.variables['pm_value'].set(0)
            self.variables['mpm_value'].set(0)

    def update_ch_code (self):
        equip_code = self.variables["equip_code"].get()

        # Check if equip_code starts with 1 or 2 and has 4 characters
        if len(equip_code) == 4 and equip_code[0] in ["1", "2"]:
            ch4_code = "3" + equip_code[1:] + "-4"
            ch5_code = "3" + equip_code[1:] + "-5"
            ch6_code = "3" + equip_code[1:] + "-6"
            gnd_code = "9" + equip_code[1:]
            al_code = "0" + equip_code[1:]
            # Update the value of the ch entries
            self.variables["ch4_code"].set(ch4_code)
            self.variables["ch5_code"].set(ch5_code)
            self.variables["ch6_code"].set(ch6_code)
            self.variables["gnd_code"].set(gnd_code)
            self.variables["al_code"].set(al_code)

    def update_chTP_code (self):
        tp_code = self.variables["tp_code"].get()

        # Check if equip_code starts with 1 or 2 and has 4 characters
        if len(tp_code) == 4 and tp_code[0] in ["8"]:
            chTP_code = "4" + tp_code[1:]
            # Update the value of the ch entries
            self.variables["chTP_code"].set(chTP_code)

    def update_tp_code(self, event=None):
        """Update the TP code based on the selected voltage."""

        # Get the current TP code
        current_tp_code = self.variables['tp_code'].get()

        # Get the selected voltage
        selected_voltage = self.voltage_value.get()

        # Check if the current TP code meets the conditions
        if len(current_tp_code) == 4 and current_tp_code.startswith('8') and current_tp_code[1] in ['1', '2', '3', '9']:
            # Update the TP code based on the selected voltage
            valid = bool(0)
            if selected_voltage == '138KV':
                new_tp_code = '83B1'
                valid = bool(1)
            elif selected_voltage == '69KV':
                new_tp_code = '82B1'
                valid = bool(1)
            elif selected_voltage == '34,5KV':
                new_tp_code = '89B1'
                valid = bool(1)
            elif selected_voltage == '23KV':
                new_tp_code = '81B1'
                valid = bool(1)
            elif selected_voltage == '13,8KV':
                new_tp_code = '81B1'
                valid = bool(1)

            # Update the TP code entry
            if valid:
                self.variables['tp_code'].set(new_tp_code)

    def update_tc_code(self):
         # Enable the "Tensão da SE" entry widget
        self.entries["se_voltage"].configure(state="normal")

        # Get the selected voltages from the checkboxes
        selected_voltages = [voltage.replace("KV", "").replace(",", ".") for voltage, var in self.checkbox_vars.items() if var.get()]

        # Set the value of the "Tensão da SE" entry widget
        if selected_voltages:
            self.variables['se_voltage'].set('/'.join(selected_voltages) + "KV")
        else:
            self.variables['se_voltage'].set("")

        # Disable the "Tensão da SE" entry widget after setting its value
        self.entries["se_voltage"].configure(state="disabled")


    def update_voltage_entry(self):
        # Enable the "Tensão da SE" entry widget
        self.entries["se_voltage"].configure(state="normal")

        # Get the selected voltages from the checkboxes
        selected_voltages = [voltage.replace("KV", "").replace(",", ".") for voltage, var in self.checkbox_vars.items() if var.get()]

        # Set the value of the "Tensão da SE" entry widget
        if selected_voltages:
            self.variables['se_voltage'].set('/'.join(selected_voltages) + "KV")
        else:
            self.variables['se_voltage'].set("")

        # Disable the "Tensão da SE" entry widget after setting its value
        self.entries["se_voltage"].configure(state="disabled")

    def update_tc_code(self, *args):
        equip_code = self.variables["equip_code"].get()

        # Check if equip_code starts with 1 or 2 and has 4 characters
        if len(equip_code) == 4 and equip_code[0] in ["1", "2"]:
            # Check the value of estado_entry
            if self.estado_value.get() == "Piauí":
                tc_code = "9" + equip_code[1:]
            else:
                tc_code = "7" + equip_code[1:]
        else:
            # If the conditions are not met, keep the current value of tc_code
            tc_code = self.variables["tc_code"].get()

        # Update the value of the tc_code entry
        self.variables["tc_code"].set(tc_code)

    def replace_values(self):
        
        # Extract the state
        estado = self.estado_value.get()
        
        # Create the EQ_name variable
        EQ_name = f"EQUATORIAL {estado}"

         # Remove "KV" from the voltage value
        voltage_value = self.voltage_value.get().replace("KV", "")

        find_replace_dict = {
        "!**#": self.variables['ucp_value'].get().upper(),
        "!**PP": self.variables['pp_value'].get().upper(),
        "!**!": self.variables['djpp_value'].get().upper(),
        "!**/": self.variables['um_value'].get().upper(),
        "!**\"": self.variables['mpm_value'].get().upper(),
        "!**$": self.variables['se_name'].get().upper(),
        "!**&": self.variables['se_code'].get().upper(),
        "!**?": voltage_value,
        "!**@": self.variables['equip_code'].get().upper(),
        "!**TC": self.variables['tc_code'].get().upper(),
        "!**TP": self.variables['tp_code'].get().upper(),
        "!**ME": self.variables['mpm_value'].get().upper(),
        "!**PJ": self.variables['pj_name'].get().upper(),
        "!**PN": self.variables['pj_fname'].get().upper(),
        "!**DT": self.variables['date'].get().upper(),
        "!**TEN": self.variables['se_voltage'].get().upper(),
        "!**EQ": EQ_name.upper(),
        "31!**CH-4": self.variables['ch4_code'].get().upper(),
        "31!**CH-5": self.variables['ch5_code'].get().upper(),
        "31!**CH-6": self.variables['ch6_code'].get().upper(),
        "91!**CH": self.variables['gnd_code'].get().upper(),
        "01!**CH": self.variables['al_code'].get().upper(),
        "!**chTP": self.variables['chTP_code'].get().upper(),
        "!**CHTP": self.variables['chTP_code'].get().upper(),
        "!**DJCC": self.variables['DJCC_number'].get().upper(),
        "X!**CCP": self.variables['BPCC_number'].get().upper(),
        "X!**CCN": self.variables['BNCC_number'].get().upper(),
        "!**DJCA": self.variables['DJCA_number'].get().upper(),
        "X!**CAP": self.variables['BPCA_number'].get().upper(),
        "X!**CAN": self.variables['BNCA_number'].get().upper()
        
        }
        """Start a new thread to execute the replace_values method."""
        # Disable the replace button
        self.replace_button.configure(state="disabled")

        self.GDPutils.replace_cad_text(find_replace_dict)

        # Enable the replace button again after the operation completes
        self.replace_button.configure(state="normal")
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS)

    def replace_values_thread(self):
        #Do the operation in a separate thread
        threading.Thread(target=self.replace_values).start()

        
    


if __name__ == "__main__":
    app = ReplaceValuesApp()
    app.iconbitmap(default='bin/icon.ico')
    app.mainloop()

import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from tkcalendar import DateEntry
import win32evtlog
import pandas as pd
from fpdf import FPDF
from datetime import datetime

class EventLogViewer:
    def __init__(self, selected_event_ids):
        self.events = {
            "Create Service": [7030, 7045],
            "Create User": [4720, 4722, 4724, 4728],
            "Add User to Group": [4732],
            "Clear Event Log": [1102],
            "Create RDP Certificate": [1056],
            "Insert USB": [
                7030,       # Device connected
                1006,       # System event for device connection
                10000,      # USB insertion in Device Manager
                10001,      # Device setup
                20001,      # Device installed
                20002,      # Device installed successfully
                20003,      # Device installed with additional setup
                24576,      # USB insertion detected
                24577,      # USB ready to use
                24578,      # USB started
                24579,      # USB fully operational
                6416,       # New USB device connected (requires advanced auditing)
                2100,       # Device removal (general)
                2101,       # Device removal success
                2102,       # Device configured
                2103,       # Device configured successfully
                2104,       # Device removal failed
                4500,       # Volume mount for USB
                4501,       # Volume unmount for USB
                #4656,       # Access to removable storage
                4663,       # Access attempt on removable storage
                4727,       # USB insertion detected by security
                6418,       # USB device installed
                6423,       # Removable device detected
            ],
            "Disable Firewall": [2003],
            "Applocker": [8003, 8006, 8007],
            "EMET": [2],
            "Logon Failed": [4625],
            "Service Terminated Unexpectedly": [7034],
            "A service was installed in the system": [4697],
            "User Account Locked Out": [4740],
            "User Account Unlocked": [4767],
            "File Access / Deletion": [4663, 4659, 4660],
            "Terminal service session reconnected": [4778],
            "Terminal service session disconnected": [4779],
            "User Initiated Logoff": [4647],
            "A directory service object was created": [5137],
            "A directory service object was modified": [5136],
            "Permission change with old & new attributes": [4670],
            "Service Start Type Change (disable, manual, automatic)": [7040],
            "Service Start / Stop": [7036],
            "Restart Windows": [1076],
            "Shutdown Windows": [1074],
            "Logon Failure": [4625],
            "Password Change": [4723, 4724],
            "Account Disabled": [4725],
            "Account Enabled": [4731],
            "Access to Network Resource": [5140],
            "Service Failure": [7031, 7034],
            "Service Started": [7036],
            "Program Installation": [19],
            "Update Installation": [2],
            "Security Policy Change": [4739],
            "Firewall Block": [5152, 5153],
            "Driver Installation": [6006]
        }
        self.selected_event_ids = selected_event_ids
        self.logType = "Security"
        self.hand = None
        self.event_data = []

    def connect_log(self):
        try:
            self.hand = win32evtlog.OpenEventLog(None, self.logType)
            if not self.hand:
                raise Exception(f"Error al abrir el registro de eventos: {self.logType}")
            print("Conectado al registro de eventos correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el registro de eventos: {str(e)}")

    def disconnect_log(self):
        if self.hand:
            win32evtlog.CloseEventLog(self.hand)

    def read_events(self):
        self.event_data.clear()
        if not self.hand:
            messagebox.showerror("Error", "No se pudo abrir el registro de eventos.")
            return
        try:
            while True:
                events = win32evtlog.ReadEventLog(
                    self.hand,
                    win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                    0
                )
                if not events:
                    print("No se encontraron más eventos.")
                    break
                for event in events:
                    if event.EventID in self.selected_event_ids:
                        event_type = next((type for type, ids in self.events.items() if event.EventID in ids), None)
                        if event_type:
                            event_time = event.TimeGenerated
                            event_date = event_time.strftime("%d-%m-%Y")
                            event_hour = event_time.strftime("%H:%M:%S")
                            
                            self.event_data.append({
                                "Event ID": event.EventID,
                                "Date": event_date,
                                "Time": event_hour,
                                "Type": event_type
                            })
            print(f"Se leyeron {len(self.event_data)} eventos.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el registro de eventos: {e}")

class GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Registro de Eventos")
        self.root.geometry("1367x780")
        self.selected_event_ids = set()
        self.log_viewer = EventLogViewer(self.selected_event_ids)
        self.search_var = tk.StringVar()

        # Crear una pestaña para la selección de eventos
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)

        # Frame principal para el visualizador de eventos
        main_frame = tk.Frame(notebook)
        notebook.add(main_frame, text="Visualizador de Eventos")


        # Frame de selección de eventos con scrollbar
        selection_frame = tk.Frame(notebook)
        notebook.add(selection_frame, text="Seleccionar Eventos")

        canvas = tk.Canvas(selection_frame)
        scrollbar = tk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        def on_mouse_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mouse_wheel)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for category, ids in self.log_viewer.events.items():
            category_label = tk.Label(scrollable_frame, text=category, font=("Arial", 10, "bold"))
            category_label.pack(anchor="w", pady=(5, 0))
            for event_id in ids:
                var = tk.BooleanVar(value=True)
                checkbutton = tk.Checkbutton(
                    scrollable_frame,
                    text=f"ID: {event_id} - {category}",
                    variable=var,
                    command=lambda eid=event_id, var=var: self.toggle_event_id(eid, var)
                )
                checkbutton.pack(anchor="w")
                self.selected_event_ids.add(event_id)

        self.create_main_ui(main_frame)

    def create_main_ui(self, main_frame):
        # Barra de búsqueda
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill="x", padx=10, pady=10)
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'), foreground='#000080')
        style.configure("Treeview", font=('Helvetica', 9), rowheight=25)
        
        tk.Button(top_frame, text="Cargar Eventos", command=self.load_events, bg="#FFA500", fg="white").pack(side="left", padx=5)
        
        
        start_date_label = tk.Label(top_frame, text="Inicio:", fg="#00008B", font=("Arial", 11,"bold"))
        start_date_label.pack(side="left", padx=5)
        self.start_date_entry = DateEntry(top_frame, width=11, background="#00008B", foreground="white", borderwidth=2)
        self.start_date_entry.pack(side="left", padx=5)

        end_date_label = tk.Label(top_frame, text="Fin:",fg="#00008B", font=("Arial", 11,"bold"))
        end_date_label.pack(side="left", padx=5)
        self.end_date_entry = DateEntry(top_frame, width=11, background="#00008B", foreground="white", borderwidth=2)
        self.end_date_entry.pack(side="left", padx=5)

        tk.Button(top_frame, text="Filtrar Eventos", command=self.update_treeview, bg="#1E90FF", fg="white").pack(side="left", padx=5)

        search_label = tk.Label(top_frame, text="Buscar:", fg="#E9967A", font=("Arial", 11,"bold"))
        search_label.pack(side="left", padx=5)

        search_entry = tk.Entry(top_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=5)
        search_entry.insert(0, "")

        exportar_label = tk.Label(top_frame, text="Exportar:", fg="green", font=("Arial", 11,"bold"))
        exportar_label.pack(side="left", padx=5)

        tk.Button(top_frame, text="Formato TXT", command=self.export_to_txt, bg="#98FB98", fg="black").pack(side="left", padx=5)
        tk.Button(top_frame, text="Formato PDF", command=self.export_to_pdf, bg="#98FB98", fg="black").pack(side="left", padx=5)
        tk.Button(top_frame, text="Formato EXCEL", command=self.export_to_excel, bg="#98FB98", fg="black").pack(side="left", padx=5)

        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        tree_scroll = tk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll.pack(side="right", fill="y")
        
        columns = ("Index", "Event ID", "Date", "Time", "Type")
        self.event_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set, style="Treeview")

        self.event_tree.heading("Index", text="N°")
        self.event_tree.heading("Event ID", text="Event ID")
        self.event_tree.heading("Date", text="Fecha")
        self.event_tree.heading("Time", text="Hora")
        self.event_tree.heading("Type", text="Tipo de Evento")
        self.event_tree.pack(fill="both", expand=True)

        self.event_tree.column("Index", width=50, anchor="center")
        self.event_tree.column("Event ID", width=100, anchor="w")
        self.event_tree.column("Date", width=100, anchor="center")
        self.event_tree.column("Time", width=100, anchor="center")
        self.event_tree.column("Type", width=800, anchor="w")

        tree_scroll.config(command=self.event_tree.yview)

        self.filtered_data = []
        self.search_var.trace("w", self.update_treeview)

    def load_events(self):
        self.log_viewer.connect_log()
        self.log_viewer.read_events()
        self.log_viewer.disconnect_log()
        self.update_treeview()

    def update_treeview(self, *args):
        search_text = self.search_var.get().lower()
        start_date = self.start_date_entry.get_date()  # Obtén la fecha de inicio
        end_date = self.end_date_entry.get_date()  # Obtén la fecha de fin
        
        self.filtered_data = [
            event for event in self.log_viewer.event_data 
            if (search_text in f"{event['Event ID']} {event['Date']} {event['Time']} {event['Type']}".lower()) and
            (start_date <= datetime.strptime(event['Date'], "%d-%m-%Y").date() <= end_date)
        ]
        
        # Limpiar el treeview y agregar los eventos filtrados
        for i in self.event_tree.get_children():
            self.event_tree.delete(i)
        for index, event in enumerate(self.filtered_data, start=1):
            self.event_tree.insert("", "end", values=(index, event['Event ID'], event['Date'], event['Time'], event['Type']))

    def toggle_event_id(self, event_id, var):
        if var.get():  # Si está marcado, agregarlo
            self.selected_event_ids.add(event_id)
        else:  # Si está desmarcado, eliminarlo
            self.selected_event_ids.discard(event_id)
        self.update_treeview()  # Actualiza la vista con los eventos filtrados

    def export_to_txt(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, 'w') as file:
                file.write("Registro de Eventos Sospechosos\n\n")
                for index, event in enumerate(self.filtered_data, start=1):
                    file.write(f"N°: {index} | ID: {event['Event ID']} | Fecha: {event['Date']} | Hora: {event['Time']} | Tipo: {event['Type']}\n")
            messagebox.showinfo("Exportación", "Los datos han sido exportados a TXT correctamente.")

    def export_to_pdf(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=11)
            pdf.cell(0, 10, "Registro de Eventos Sospechosos", ln=True)
            pdf.ln(10)
            for index, event in enumerate(self.filtered_data, start=1):
                pdf.cell(0, 10, f"N°: {index} | ID: {event['Event ID']} | Fecha: {event['Date']} | Hora: {event['Time']} | Tipo: {event['Type']}", ln=True)
            pdf.output(file_path)
            messagebox.showinfo("Exportación", "Los datos han sido exportados a PDF correctamente.")

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            data_frame = pd.DataFrame(self.filtered_data)
            data_frame.index += 1
            data_frame.columns = ["Event ID", "Fecha", "Hora", "Tipo"]
            data_frame.index.name = "N°"
            data_frame.to_excel(file_path)
            messagebox.showinfo("Exportación", "Los datos han sido exportados a Excel correctamente.")

if __name__ == "__main__":
    root = tk.Tk()
    app = GUI(root)
    root.mainloop()

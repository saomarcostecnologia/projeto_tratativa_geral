import customtkinter as ctk
from tkinter import messagebox
from use_cases import FileTreatmentUseCase


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.appearance()
        self.use_case = FileTreatmentUseCase()
        self.todo_sistema()

    def layout_config(self):
        self.title_var = ctk.StringVar()  # Variável de controle para o título do layout
        self.title_var.set("Sistema de Tratamento Equipe Contabilidade - Novonor")
        self.title(self.title_var.get())  # Define o título inicial
        self.geometry("700x500")
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x_offset = (self.winfo_screenwidth() - width) // 2
        y_offset = (self.winfo_screenheight() - height) // 2
        self.geometry(f"{width}x{height}+{x_offset}+{y_offset}")

    def appearance(self):
        self.lb_apm = ctk.CTkLabel(
            self, text="Tema", bg_color="transparent", text_color=["#000", "#fff"]
        ).place(x=50, y=440)
        self.opt_apm = ctk.CTkOptionMenu(
            self, values=["Light", "Dark", "System"], command=self.change_apm
        ).place(x=50, y=465)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)

    def select_file(self):
        file_path = self.use_case.select_file()
        if file_path:
            self.selected_file_label = ctk.CTkLabel(
                self,
                text="Arquivo selecionado: " + file_path.file_path,
                font=("Century Gothic", 12),
            ).place(x=50, y=150)
            self.file_path = file_path.file_path

    def todo_sistema(self):
        frame = ctk.CTkFrame(
            self,
            width=700,
            height=50,
            corner_radius=0,
            bg_color="teal",
            fg_color="teal",
        ).place(x=0, y=10)

        title = ctk.CTkLabel(
            frame,
            text="Sistema de Tratamento - Novonor",
            font=("Century Gothic", 22),
            text_color="#fff",
            bg_color="teal",
        ).place(x=35, y=25)

        btn_select_file = ctk.CTkButton(
            self, text="Selecionar Arquivo", command=self.select_file
        ).place(x=45, y=110)

        span = ctk.CTkLabel(
            self,
            text="Por Favor, selecione o arquivo para tratamento",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        ).place(x=50, y=70)

        self.mode_box = ctk.CTkComboBox(
            self,
            values=["Salário", "Balancete", "Razão"],
            font=("Century Gothic", 14),
            dropdown_font=("Century Gothic", 14),
            state="readonly",  # Configurando o combobox para ser somente leitura
            corner_radius=20,
        )
        self.mode_box.set("Salário")
        self.mode_box.place(x=550, y=70)

        btn_execute = ctk.CTkButton(
            self,
            text="Executar Tratamento".upper(),
            command=self.submit,
            fg_color="#151",
            hover_color="#131",
        ).place(x=520, y=465)

        btn_clear_selection = ctk.CTkButton(
            self,
            text="Limpar Seleção".upper(),
            command=self.clear_file_selection,
            fg_color="#555",
            hover_color="#333",
        ).place(x=355, y=465)

    def submit(self):
        if hasattr(self, "file_path") and self.file_path:
            selected_mode = self.mode_box.get()
            if selected_mode == "Salário":
                self.tratar_salario()
            elif selected_mode == "Balancete":
                self.tratar_balancete()
            elif selected_mode == "Razão":
                self.tratar_razao()
            else:
                print("Modo inválido selecionado.")
        else:
            messagebox.showerror(
                "Erro",
                "Nenhum arquivo selecionado. Selecione um arquivo antes de executar o tratamento.",
            )

    def clear_file_selection(self):
        if hasattr(self, "selected_file_label"):
            ctk.CTkLabel(
                self,
                text="Nenhum arquivo selecionado." + (" ") * 200,
                font=("Century Gothic", 12),
            ).place(x=50, y=150)
        if hasattr(self, "file_path"):
            self.file_path = None
        messagebox.showinfo("Limpeza Concluída", "Seleção de arquivo limpa!")

    def tratar_salario(self):
        success = self.use_case.tratar_salario(self.file_path)
        if success:
            self.clear_file_selection()
    
    def tratar_balancete(self):
        sucess = self.use_case.tratar_balancete(self.file_path)
        if sucess:
            self.clear_file_selection()

    def tratar_razao(self):
        sucess = self.use_case.tratar_razao(self.file_path)
        if sucess:
            self.clear_file_selection()


if __name__ == "__main__":
    app = App()
    app.mainloop()

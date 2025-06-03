import os
import shutil
import pyodbc
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

class AccessExplodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Access Database Exploder")
        self.root.geometry("800x600")
        
        # Variáveis para armazenar os caminhos dos arquivos
        self.input_files = []
        self.output_dir = tk.StringVar()
        
        # Criar interface
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Área de seleção de arquivos de entrada
        ttk.Label(main_frame, text="Arquivos de Entrada (.mdb):").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        # Lista de arquivos selecionados
        self.files_listbox = tk.Listbox(main_frame, height=5, width=70)
        self.files_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Botões para adicionar/remover arquivos
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        ttk.Button(btn_frame, text="Adicionar Arquivo", command=self.add_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remover Selecionado", command=self.remove_file).pack(side=tk.LEFT)
        
        # Seleção do diretório de saída
        ttk.Label(main_frame, text="Diretório de Saída:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=60).grid(row=4, column=0, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Selecionar", command=self.select_output_dir).grid(row=4, column=1, padx=5)
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # Botão processar
        ttk.Button(main_frame, text="Processar", command=self.process_files).grid(row=7, column=0, columnspan=2, pady=10)
        
    def add_file(self):
        if len(self.input_files) >= 2:
            messagebox.showwarning("Aviso", "Máximo de 2 arquivos permitidos!")
            return
            
        file_path = filedialog.askopenfilename(
            filetypes=[("Access Database", "*.mdb")],
            title="Selecione um arquivo Access"
        )
        
        if file_path and file_path not in self.input_files:
            self.input_files.append(file_path)
            self.files_listbox.insert(tk.END, file_path)
            
    def remove_file(self):
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.input_files.pop(index)
            self.files_listbox.delete(index)
            
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Selecione o diretório de saída")
        if directory:
            self.output_dir.set(directory)
            
    def process_files(self):
        if not self.input_files:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo de entrada!")
            return
            
        if not self.output_dir.get():
            messagebox.showerror("Erro", "Selecione um diretório de saída!")
            return
            
        self.progress['value'] = 0
        total_steps = len(self.input_files)
        step_size = 100 / total_steps
        
        for input_file in self.input_files:
            try:
                # Criar diretório de saída específico para cada arquivo
                input_name = Path(input_file).stem
                output_specific_dir = os.path.join(self.output_dir.get(), f"Exploded_{input_name}")
                os.makedirs(output_specific_dir, exist_ok=True)
                
                self.status_label['text'] = f"Processando: {input_name}"
                self.root.update()
                
                # Chamar a função de processamento
                self.explodir_access(input_file, output_specific_dir)
                
                self.progress['value'] += step_size
                self.root.update()
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar {input_file}: {str(e)}")
                
        self.status_label['text'] = "Processamento concluído!"
        messagebox.showinfo("Sucesso", "Processamento finalizado com sucesso!")
        
    def get_connection_string(self, db_path):
        return (
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
            f"DBQ={db_path};"
        )
        
    def listar_tabelas(self, conn):
        cursor = conn.cursor()
        tabelas = []
        for row in cursor.tables(tableType='TABLE'):
            tabelas.append(row.table_name)
        return tabelas
        
    def contar_registros(self, conn, tabela):
        cursor = conn.cursor()
        try:
            cursor.execute(f"SELECT COUNT(*) FROM [{tabela}]")
            return cursor.fetchone()[0]
        except Exception as e:
            print(f"Erro ao contar registros na tabela {tabela}: {e}")
            return 0
            
    def copiar_banco_origem(self, input_path, destino_path):
        try:
            shutil.copy(input_path, destino_path)
            return True
        except Exception as e:
            print(f"Erro ao copiar o banco de origem para {destino_path}: {e}")
            return False
            
    def explodir_access(self, input_path, output_dir):
        try:
            # Conectar ao banco de entrada
            conn = pyodbc.connect(self.get_connection_string(input_path))
            tabelas = self.listar_tabelas(conn)
            
            # Filtrar tabelas com dados
            tabelas_com_dados = [tabela for tabela in tabelas if self.contar_registros(conn, tabela) > 0]
            
            # Processar tabelas com dados
            for tabela in tabelas_com_dados:
                self.status_label['text'] = f"Processando tabela: {tabela}"
                self.root.update()
                
                # Criar uma cópia do banco de dados original
                arquivo_saida = os.path.join(output_dir, f"{tabela}.mdb")
                if not self.copiar_banco_origem(input_path, arquivo_saida):
                    continue
                    
                try:
                    # Conectar à cópia do banco
                    conn_saida = pyodbc.connect(self.get_connection_string(arquivo_saida))
                    cursor_saida = conn_saida.cursor()
                    
                    # Excluir dados das outras tabelas
                    for outra_tabela in tabelas:
                        if outra_tabela != tabela:
                            try:
                                cursor_saida.execute(f"DELETE FROM [{outra_tabela}]")
                                conn_saida.commit()
                            except Exception:
                                pass
                                
                    # Manter apenas os dados filtrados da tabela selecionada
                    try:
                        cursor_saida.execute(
                            f"DELETE FROM [{tabela}] WHERE Importado <> 0 OR ImportadoRepetido <> 0 OR Removido <> 0 OR Temporario <> 0"
                        )
                        conn_saida.commit()
                    except Exception:
                        pass
                        
                except Exception as e:
                    print(f"Erro ao processar a tabela {tabela}: {e}")
                finally:
                    conn_saida.close()
                    
        except Exception as e:
            raise Exception(f"Erro durante o processamento: {e}")
        finally:
            conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = AccessExplodeApp(root)
    root.mainloop() 
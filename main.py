import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog

def find_excel_files():
    """Find all Excel and CSV files in the current directory."""
    files = []
    for file in os.listdir('.'):
        if file.endswith(('.xlsx', '.xls', '.csv')):
            files.append(file)
    return files

def select_file(files):
    """Display a simple GUI for file selection."""
    if not files:
        messagebox.showerror("错误", "当前目录下没有找到Excel或CSV文件！")
        return None
    
    root = tk.Tk()
    root.title("选择文件")
    root.geometry("400x300")
    
    tk.Label(root, text="请选择要处理的文件：", pady=10).pack()
    
    var = tk.StringVar(root)
    
    for i, file in enumerate(files):
        tk.Radiobutton(root, text=f"{i+1}. {file}", variable=var, value=file, anchor="w").pack(fill="x", padx=20)
    
    selected_file = [None]  # Use a list to store the result
    
    def on_submit():
        selected_file[0] = var.get()
        root.destroy()
    
    tk.Button(root, text="确定", command=on_submit, pady=5).pack(pady=20)
    
    root.mainloop()
    return selected_file[0]

def select_column(columns):
    """Let user select which column to use as identifier."""
    root = tk.Tk()
    root.title("选择标识列")
    root.geometry("500x400")
    
    tk.Label(root, text="请选择用于标识回答者的列：", pady=10).pack()
    
    var = tk.StringVar(root)
    
    for i, col in enumerate(columns):
        tk.Radiobutton(root, text=f"{i+1}. {col}", variable=var, value=col, anchor="w").pack(fill="x", padx=20)
    
    selected_column = [None]  # Use a list to store the result
    
    def on_submit():
        selected_column[0] = var.get()
        root.destroy()
    
    tk.Button(root, text="确定", command=on_submit, pady=5).pack(pady=20)
    
    root.mainloop()
    return selected_column[0]

def process_file(file_path, id_column):
    """Process the selected file and generate summary."""
    try:
        # Determine file type
        if file_path.endswith('.csv'):
            # Try different encodings
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030']
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
        else:
            df = pd.read_excel(file_path)
        
        # Clean up the dataframe - remove empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Find questions - any column after "线路感想" or containing "感想"
        感想_columns = [col for col in df.columns if '感想' in col]
        if not 感想_columns:
            messagebox.showerror("错误", "未找到包含'感想'的列！")
            return None
            
        感想_column = 感想_columns[0]
        question_index = list(df.columns).index(感想_column)
        question_columns = [感想_column] + list(df.columns[question_index+1:])
        
        # Create summary
        summary = []
        for question in question_columns:
            section = [f"{question}\n"]
            
            for _, row in df.iterrows():
                # Skip if no identifier or no answer
                if pd.isna(row[id_column]) or pd.isna(row[question]) or \
                    str(row[question]).strip() in ('', '(空)'):
                    continue
                
                identifier = row[id_column]
                answer = row[question]
                section.append(f"- {identifier}: {answer}")
            
            section.append("\n")  # Add a blank line between sections
            summary.extend(section)
        
        return summary
    
    except Exception as e:
        messagebox.showerror("处理错误", f"处理文件时出错: {str(e)}")
        return None

def save_summary(summary, original_file):
    """Save the summary to a text file."""
    base_name = os.path.splitext(original_file)[0]
    output_file = f"{base_name}_整理结果.txt"
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(summary))
    
    return output_file

def main():
    # Create a simple splash screen
    root = tk.Tk()
    root.title("问卷回答整理工具")
    root.geometry("400x200")
    
    tk.Label(root, text="问卷回答整理工具", font=("Helvetica", 16)).pack(pady=20)
    tk.Label(root, text="正在搜索Excel和CSV文件...").pack()
    
    # Update the UI
    root.update()
    
    # Find files
    files = find_excel_files()
    
    # Close splash
    root.destroy()
    
    # Select file
    selected_file = select_file(files)
    if not selected_file:
        return
    
    # Read file to get columns
    try:
        if selected_file.endswith('.csv'):
            # Try different encodings
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030']
            for encoding in encodings:
                try:
                    df = pd.read_csv(selected_file, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
        else:
            df = pd.read_excel(selected_file)
        
        # Clean up the dataframe - remove empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
    except Exception as e:
        messagebox.showerror("读取错误", f"读取文件时出错: {str(e)}")
        return
    
    # Select identifier column
    id_column = select_column(df.columns[:10])
    if not id_column:
        return
    
    # Process file
    summary = process_file(selected_file, id_column)
    if not summary:
        return
    
    # Save summary
    output_file = save_summary(summary, selected_file)
    
    # Show completion message
    messagebox.showinfo("处理完成", f"处理完成！\n结果已保存至: {output_file}")

if __name__ == "__main__":
    main()
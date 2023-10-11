#!/usr/bin/python3

from pathlib import Path
from tkinter import Tk, messagebox, filedialog, Label, Button, Entry, Frame, StringVar

import docx2txt


class DocxImageExtractorApp:
    dir_path_var: StringVar

    def __init__(self):
        root = Tk()
        root.title('DOCX image extractor')
        root.resizable(False, False)
        root.geometry('600x120')

        main_frame = Frame(root)
        main_frame.pack(pady=20)

        Label(main_frame, text='Путь к папке с docx-файлами:').grid(row=0, column=0)
        self.dir_path_var = StringVar()
        Entry(main_frame, width=50, textvariable=self.dir_path_var).grid(row=0, column=1, padx=5)
        Button(main_frame, text='...', command=self.select_dir_callback).grid(row=0, column=2)

        button_frame = Frame(root)
        button_frame.pack()

        Button(button_frame, text='Распаковать', command=self.extract_images_callback).grid(row=0, column=1)
        Button(button_frame, text='Выход', command=root.destroy).grid(row=0, column=2)

        root.mainloop()

    def select_dir_callback(self) -> None:
        dir_path = filedialog.askdirectory(title='Выберите папку с DOCX-файлами')
        if not not dir_path:
            self.dir_path_var.set(dir_path)

    def extract_images_callback(self) -> None:
        docx_files_path = Path(self.dir_path_var.get())

        if not docx_files_path.is_dir():
            messagebox.showwarning('Внимание!', f'Папка "{docx_files_path}" не существует')
            return

        file_count = self.extract_images(docx_files_path)

        if file_count == 0:
            messagebox.showwarning('Внимание!', f'DOCX-файлы не обнаружены в папке: "{docx_files_path}"')
            return

        messagebox.showinfo('Готово', f'Изображения из ({file_count}) файлов извлечены')

    @staticmethod
    def extract_images(dir_path: Path) -> int:
        print('dir_path:', dir_path)

        file_count = 0

        for file_path in dir_path.iterdir():
            if file_path.is_file() and file_path.suffix == '.docx':
                file_count += 1
                print('- Extracting images from:', file_path)
                images_path = dir_path / (file_path.name + '_images')
                images_path.mkdir(exist_ok=True)
                docx2txt.process(file_path, images_path)

        return file_count


if __name__ == '__main__':
    DocxImageExtractorApp()

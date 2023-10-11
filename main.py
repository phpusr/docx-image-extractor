from pathlib import Path
from tkinter import Tk, Label, Button, messagebox, filedialog, Entry, END

import docx2txt


class DocxImageExtractorApp:
    file_path_text: Entry

    def __init__(self):
        root = Tk()
        root.title('DOCX image extractor')
        root.resizable(False, False)
        root.geometry('600x130')

        label = Label(root, text='Путь к папке с docx-файлами:')
        label.grid(row=0, column=0, pady=20)

        self.file_path_text = Entry(root, width=50)
        self.file_path_text.grid(row=0, column=1, padx=5)

        select_dir_button = Button(root, text='...', command=self.select_dir_callback)
        select_dir_button.grid(row=0, column=2)

        extract_button = Button(root, text='Распаковать', command=self.extract_images_callback)
        extract_button.grid(row=1, column=0, columnspan=3)

        root.mainloop()

    def select_dir_callback(self):
        self.file_path_text.delete(0, END)
        self.file_path_text.insert(0, filedialog.askdirectory(title='Выберите папку с DOCX-файлами'))

    def extract_images_callback(self):
        docx_files_path = Path(self.file_path_text.get())
        self.extract_images(docx_files_path)
        messagebox.showinfo('', 'Изображения из файлов извлечены')

    @staticmethod
    def extract_images(dir_path: Path):
        print('dir_path:', dir_path)

        for file_path in dir_path.iterdir():
            if file_path.is_file() and file_path.suffix == '.docx':
                print('- Extracting images from:', file_path)
                images_path = dir_path / (file_path.name + '_images')
                images_path.mkdir(exist_ok=True)
                docx2txt.process(file_path, images_path)


if __name__ == '__main__':
    DocxImageExtractorApp()

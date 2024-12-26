from tkinter import *
from spire.doc import *
from tkinter import filedialog, messagebox
from tkinter.simpledialog import askstring, askfloat
from spire.doc import Color
window = Tk()
window.title("Word Document Editor")

w = Document()
word_file = None

def choosfile():
    global w, word_file
    word_file = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
    if word_file:  # Check if a file was selected
        w.LoadFromFile(word_file)

button = Button(window, text='Choose File', bg='red', command=choosfile)
button.pack()

def write_word():
    global w
    if not word_file:
        messagebox.showwarning("Warning", "No file loaded. Please load a file first.")
        return

    s1 = w.AddSection()
    s3 = s1.AddParagraph()
    sentence = askstring("Enter Sentence", "Please enter a sentence:")
    if sentence:  # Check if the user entered something
        s3.AppendText(sentence)
        s3.Format.HorizontalAlignment = HorizontalAlignment.Distribute
        
        new_location = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if new_location:  # Check if a save location was chosen
            w.SaveToFile(new_location)

button = Button(window, text='Write Sentence', bg='blue', command=write_word)
button.pack()

def add_text_to_new_word():
    text = askstring("Input Text", "Please enter the text you want to add:")
    if text:
        new_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if new_file:
            new_doc = Document()
            section = new_doc.Sections.get_Item(0)
            pargh = section.AddParagraph()
            pargh.AppendText(text)
            new_doc.SaveToFile(new_file)  # ذخیره سند جدید
            messagebox.showinfo("Text Added", "Text has been added to the new document.")


add_text_to_new_word_button = Button(window, text='add_text_to_new_word', bg='purple', command=add_text_to_new_word)
add_text_to_new_word_button.pack()

def photo_on_word():
    global w
    if not word_file:
        messagebox.showwarning("Warning", "No file loaded. Please load a file first.")
        return

    text = w.AddSection()
    p = text.AddParagraph()
    image = filedialog.askopenfilename(filetypes=[("Image", '*.jpg')])
    if image:  # Check if an image file was selected
        Width_photo = askfloat("Enter Width", "Please enter a width:")
        Height_photo = askfloat("Enter Height", "Please enter a height:")
        if Width_photo and Height_photo:  # Check if both dimensions were entered
            pic = p.AppendPicture(image)
            pic.Width = Width_photo
            pic.Height = Height_photo
            pic.TextWrappingStyle = TextWrappingStyle.Square
            messagebox.showinfo("Image Added", "Image has been added to the document.")

photo_button = Button(window, text='Add Photo to Word', bg='yellow', command=photo_on_word)
photo_button.pack()

def file_on_file():
    if not word_file:
        messagebox.showwarning("Warning", "No file loaded. Please load a file first.")
        return

    file2 = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
    if file2:  # Check if a second file was selected
        new_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if new_file:  # Check if a save location was chosen
            w.LoadFromFile(word_file)
            w.InsertTextFromFile(file2, FileFormat.Auto)
            w.SaveToFile(new_file)

file_button = Button(window, text='Insert File into Document', bg='red', command=file_on_file)
file_button.pack()

def Save_file():
    global word_file, w
    if word_file:
        w.SaveToFile(word_file)
        messagebox.showinfo("File Saved", "File has been saved successfully.")
    else:
        new_file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if new_file_path:  # Ensure the user selected a file
            w.SaveToFile(new_file_path)
            word_file = new_file_path
            messagebox.showinfo("File Saved", "File has been saved successfully.")

button = Button(window, text='Save File', bg='green', command=Save_file)
button.pack()

window.mainloop()
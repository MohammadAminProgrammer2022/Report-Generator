from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import Canvas,  Button, PhotoImage, \
    Frame, Label, messagebox, StringVar
from tkinter.ttk import Progressbar

from pathlib import Path
import sys
import threading

from brain import DataReplacer


def resource_path(relative_path):
    try:
        # When running as a bundled .exe
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        # When running as a script
        base_path = Path(__file__).parent
    return base_path / relative_path


logo_path = resource_path("./software.ico")
template_path = resource_path("assets/template.docx")
image_1_png = resource_path("./assets/frame0/image_1.png")
btn = resource_path("./assets/frame0/button_1.png")

# Initializing DataReplacer class
dr = DataReplacer()


# drop trigger function | days file
def day_drop(event):
    '''
    This function triggers the event
    of dropping a file and extracts its
    path and then check if it's readable or not.
    '''
    file_path = event.data
    is_done, msg = dr.get_days(file_path)
    if is_done:
        correct_input(msg)
        return 1

    show_error(msg)
    return 0


# drop trigger function | students' data file
def data_drop(event):
    '''
    This function triggers the event
    of dropping a file and extracts its
    path and then check if it's readable or not.
    '''
    file_path = event.data
    is_done, msg = dr.get_user_data(file_path)
    if is_done:
        correct_input(msg)
        return 1
    
    show_error(msg)
    return 0


# processing all the dropped data
def process_data():
    if not dr.days_file:
        show_error('لطفا فایل روزها را وارد کرده و مجددا اقدام کنید')
        return 0
    if not dr.data_file:
        show_error('لطفا فایل اطلاعات حفاظ را وارد کرده و مجددا اقدام کنید')
        return 0
    if dr.counter == 1:
        show_error('شما برای بار دوم اقدام به تهیه گزارش کردید. برای این کار برنامه را مجددا راه‌اندازی کنید و گزارشات تهیه شده را بردارید.')
        return 0
    progress_bar["value"] = 0
    progress_var.set("0%")
    progress_bar.update()
    def processing_data():
        total_rows = dr.total_rows
        i = 1
        for _ in dr.data_replace_main(template_path):
            value2show = int(round(((i / total_rows) * 100), 2))
            progress_bar["value"] = value2show
            progress_var.set(f"{value2show}%")
            progress_bar.update()
            i += 1
        print('process done!')
    threading.Thread(target=processing_data).start()
    messagebox.showinfo('اتمام پردازش', 'گزارشات با موفقیت تهیه شدند')


def start_process():
    print('start bcuz clicked!')
    try:
        process_data()
        return 1
    except Exception as e:
        messagebox.showerror('خطا تهیه گزارش', 'خطای سیستم:\n{e}')
        return 0


def show_error(msg):
    messagebox.showerror("خطا قالب فایل", msg)


def correct_input(msg):
    messagebox.showinfo("دریافت فایل", msg)


window = TkinterDnD.Tk()

window.geometry("800x500")
window.configure(bg = "#FFFFFF")


canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 500,
    width = 800,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
image_image_1 = PhotoImage(
    file= image_1_png)
image_1 = canvas.create_image(
    400.0,
    250.0,
    image=image_image_1
)

button_image_1 = PhotoImage(
    file= btn)
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=start_process,
    relief="flat"
)
button_1.place(
    x=484.0,
    y=425.0,
    width=261.0,
    height=72.0
)

canvas.create_text(
    449.0,
    11.0,
    anchor="nw",
    text="نرم‌افزار تهیه گزارشات حفظ",
    fill="#FFFBFB",
    font=("ZillaSlab Regular", 32 * -1)
)

canvas.create_text(
    443.0,
    61.0,
    anchor="nw",
    text="ویژه حفاظ و  قرآن آموزان موسسه بینات",
    fill="#FFFBFB",
    font=("ZillaSlab Regular", 18 * -1)
)

canvas.create_text(
    629.0,
    111.0,
    anchor="nw",
    text="محل قرار دادن فایل روزها",
    fill="#FFFFFF",
    font=("Inter", 16 * -1)
)

canvas.create_text(
    571.0,
    238.0,
    anchor="nw",
    text="محل قرار دادن فایل اطلاعات حفاظ",
    fill="#FFFFFF",
    font=("Inter", 16 * -1)
)


# Drop zone | days.xlsx
day_drop_frame = Frame(window, bg="lightgray", bd=2, relief="ridge", width=380, height=150)
day_drop_frame.place(x=446, y=136, width=335, height=80)
day_drop_frame.pack_propagate(False)

day_drop_label = Label(day_drop_frame, text='فایل روز های ماه را اینجا قرار دهید', bg="lightgray")
day_drop_label.pack(expand=True)

day_drop_frame.drop_target_register(DND_FILES)
day_drop_frame.dnd_bind('<<Drop>>', day_drop)

# Drop zone | user_inputs.xlsx
data_drop_frame = Frame(window, bg="lightgray", bd=2, relief="ridge", width=380, height=150)
data_drop_frame.place(x=446, y=261, width=335, height=80)

data_drop_label = Label(data_drop_frame, text='فایل اطلاعات حفاظ را اینجا قرار دهید', bg="lightgray")
data_drop_label.pack(expand=True)

data_drop_frame.drop_target_register(DND_FILES)
data_drop_frame.dnd_bind('<<Drop>>', data_drop)

# progress bar
progress_var = StringVar(value="0%")

percent_label = Label(window, textvariable=progress_var, font=("Helvetica", 12))
percent_label.place(x=570, y=360, width=80, height=20)

progress_bar = Progressbar(window, style="green.Horizontal.TProgressbar", orient="horizontal", length=400, mode="determinate")
progress_bar.place(x=446, y=380, width=335, height=20)



window.resizable(False, False)
window.mainloop()

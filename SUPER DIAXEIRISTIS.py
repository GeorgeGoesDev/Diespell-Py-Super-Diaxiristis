# library gia sindesi me excel
import xlsxwriter

# function ypologismou tou posostou pou antistoixei ston kathe orofo gia to asanser
# vasizetai sti logiki oti o 2os orofos plironei ta dipla apo ton 1o kai o 3os ta dipla apo ton 2o k.o.k
def anelk(orofos):
    sum=0
    i=0
    epan=0
    while sum<100:
        i+=0.001
        sum=0
        epan=0
        vima=0
        for orofoi in range(0,orofos):
            vima+=i
            sum+=vima
            epan+=1
            if sum>100:
                break
    return i/100

#gia dimiourgia gui
from tkinter import *

#arxiko parathiro
root = Tk()
root.resizable(0, 0)

root.title("SUPER DIAXEIRISTIS")

#function pou kaleitai otan patithei to koumpi "ok" sto proto parathiro
def submit():
    global E1,E2,E3,E4,E5,E6,E7,E8,E9,E10,minas,etos,address,orofoi,kipouros,katharistis,asanser,suntirisi,dei_koin,petrelaio,timi_petrelaiou
    minas = E1.get()
    etos = int(E2.get())
    address = E3.get()
    orofoi=int(E4.get())
    kipouros=float(E5.get())
    katharistis=float(E6.get())
    asanser=float(E7.get())
    suntirisi=float(E8.get())
    dei_koin=float(E9.get())
    timi_petrelaiou=float(E10.get())
    diamoirasmos=anelk(orofoi)
    root.destroy()

#function pou kaleitai otan patithei to koumpi "ok" sto deftero parathiro
def submit2():
    global onoma,katanalosi,Eonoma
    global Ekatanalosi
    onoma=Eonoma.get()
    katanalosi=int(Ekatanalosi.get())
    enoikoi_window.destroy()


timi_petrelaiou=1.03048 # default timi  €/L
kipouros=0
katharistis=0
asanser=0
boiler=0
dei_koin=0
suntririsi=0

#sximatismos arxikou parathirou
L0 = Label(root, text='Καλώς ήρθατε στο SUPER ΔΙΑΧΕΙΡΙΣΤΗΣ:',bg='blue',fg='yellow')
L0.grid(row=0,columnspan=2)
L1 = Label(root, text='Μήνας:')
L1.grid(row=1,column=0, sticky=E)
E1 = Entry(root)
E1.grid(row=1,column=1)
E1.focus_set()
L2 = Label(root, text='Έτος:')
L2.grid(row=2,column=0, sticky=E)
E2 = Entry(root)
E2.grid(row=2,column=1)
L3 = Label(root, text='Διεύθυνση:')
L3.grid(row=3,column=0, sticky=E)
E3 = Entry(root)
E3.grid(row=3,column=1)
L4 = Label(root, text='Αριθμός ορόφων:')
L4.grid(row=4,column=0, sticky=E)
E4 = Entry(root)
E4.grid(row=4,column=1)
L5 = Label(root, text='Έξοδα κηπουρού:')
L5.grid(row=5,column=0, sticky=E)
E5 = Entry(root)
E5.grid(row=5,column=1)
L6 = Label(root, text='Έξοδα καθαριστή:')
L6.grid(row=6,column=0, sticky=E)
E6 = Entry(root)
E6.grid(row=6,column=1)
L7 = Label(root, text='Έξοδα ανελκυστήρα:')
L7.grid(row=7,column=0, sticky=E)
E7 = Entry(root)
E7.grid(row=7,column=1)
L8 = Label(root, text='Έξοδα συντήρησης πολυκατοικίας:')
L8.grid(row=8,column=0, sticky=E)
E8 = Entry(root)
E8.grid(row=8,column=1)
L9 = Label(root, text='ΔΕΗ κοινόχρηστων χώρων:')
L9.grid(row=9,column=0, sticky=E)
E9 = Entry(root)
E9.grid(row=9,column=1)
L10 = Label(root, text='Τιμή πετρελαίου:')
L10.grid(row=10,column=0, sticky=E)
E10 = Entry(root)
E10.grid(row=10,column=1)


button_okay = Button(root,text="Next",width=15,command=submit)
button_okay.grid(row=11,columnspan=2)

root.mainloop()

#lista pou periexei ta genika exoda
genika = (
    ['Μήνας:', minas],
    ['Έτος:', etos],
    ['Διεύθυνση:', address],
    ['Αριθμός ορόφων:', orofoi],
    ['Έξοδα κηπουρού:', kipouros],
    ['Έξοδα καθαριστή:', katharistis],
    ['Έξοδα ανελκυστήρα:', asanser],
    ['Έξοδα συντήρησης πολυκατοικίας:', suntirisi],
    ['ΔΕΗ κοινόχρηστων χώρων:', dei_koin],
    ['Τιμή πετρελαίου / λίτρο (€):', timi_petrelaiou],
    )

#dimiourgia excel me onoma to etos, to mina kai ti diefthinsi
workbook = xlsxwriter.Workbook(str(etos)+'-'+minas+'-'+address+'.xlsx')
worksheet = workbook.add_worksheet()
#dimiourgia diaforon format pou xrisimopoiountai parakato
bold = workbook.add_format({'bold': True})
money = workbook.add_format({'num_format': '#,##€'})
money.set_align('center')
money.set_align('vcenter')
cell_format = workbook.add_format()
cell_format.set_align('center')
cell_format.set_align('vcenter')
cell_format_2 = workbook.add_format()
cell_format_2.set_align('center')
cell_format_2.set_align('vcenter')
cell_format_2.set_bold()

#arxikopoiisi stilis kai grammis
row = 0
col = 0

#antigrafi tis listas genikon exodon sto excel
for i,j in (genika):
    worksheet.write(row, col, i,bold)
    if (row>3) and (row<9):
        worksheet.write(row, col + 1, j,money)
    else:
        worksheet.write(row, col + 1, j)
    row += 1

row += 1

#kaleitai i function ypologismou ton pososton gia to asanser
diamoirasmos=anelk(orofoi)

#ksexoristo parathiro gia kathe enoiko
#dimiourgountai osa kai oi orofoi
for metritis in range(0,orofoi):
    synolo=0
    eksoda_anelk=round(diamoirasmos*asanser*(metritis+1), 2)
    enoikoi_window = Tk()
    enoikoi_window.geometry("250x80")
    enoikoi_window.resizable(0, 0)
    title= str(metritis+1) + "ος όροφος"
    enoikoi_window.title(title)

    Lonoma = Label(enoikoi_window, text='Όνομα ενοίκου:')
    Lonoma.grid(row=0,column=0, sticky=E)
    Eonoma = Entry(enoikoi_window)
    Eonoma.grid(row=0,column=1)
    Eonoma.focus_set()
    orofos=metritis+1
    Lkatanalosi = Label(enoikoi_window, text='Κατανάλωση:')
    Lkatanalosi.grid(row=1,column=0, sticky=E)
    Ekatanalosi = Entry(enoikoi_window)
    Ekatanalosi.grid(row=1,column=1)
    button_okay2 = Button(enoikoi_window,text="Submit",width=15,command=submit2)
    button_okay2.grid(row=3,columnspan=2)
    button_okay2.place(relx=0.5, rely=0.8, anchor=CENTER)
    enoikoi_window.mainloop()
    xreosi_petr=katanalosi*timi_petrelaiou
    synolo=((kipouros+katharistis+suntirisi+dei_koin)/orofoi)+(eksoda_anelk+xreosi_petr)
    #lista me ta stoixeia tou enoikou
    eidika =(
        ['Όνομα ενοίκου:',onoma],
        ['Όροφος:',orofos],
        ['Έξοδα ανελκυστήρα:',eksoda_anelk],
        ['Κατανάλωση:',katanalosi],
        ['Χρέωση Πετρελαίου:',xreosi_petr],
        ['Σύνολο:',synolo],
        )

    
    #antigrafi tis listas enoikou sto excel
    for i,k in eidika:
        if metritis==0:
            worksheet.write(row,col, i, cell_format_2)
            col +=1
    row +=1
    for i in range(0,orofoi-(orofoi-1)):
        col = 0
        for k,j in eidika:
            if (col==2) or (col==4) or (col==5):
                worksheet.write(row, col, j,money)
            else:
                worksheet.write(row, col, j)
            col +=1
        row += 1 
    row -=1

#kathorismos tou megethous ton stilon
worksheet.set_column('A:A', 32)
worksheet.set_column('B:B', 25,cell_format)
worksheet.set_column('C:C', 20,cell_format)
worksheet.set_column('D:D', 13,cell_format)
worksheet.set_column('E:E', 20,cell_format)
worksheet.set_column('F:F', 8,cell_format)

#kleisimo tou excel
workbook.close()

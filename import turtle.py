n=int(input('Chọn chuỗi in ra (1,2,3): '))
Chuoi='1.HN 2.DN 3.HCM'
Chuoi1 = Chuoi[:5]
Chuoi2 = Chuoi[5:10]
Chuoi3 = Chuoi[10:]
if n==1:
    print(Chuoi1)
elif n==2:
    print(Chuoi2)
elif n==3:
    print(Chuoi3)
else:
    pass

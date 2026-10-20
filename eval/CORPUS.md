# Sentetik test verisi

Bu veri tamamen kurgusaldır. Kişiler, firmalar ve adresler gerçek değildir (`.example` alan adı). Yeniden üretmek için: `python3 eval/make_corpus.py`.

- Toplam posta: 316
- Klasörler: F-ARCH=117, F-INBOX=92, F-PROJ=42, F-SENT=65
- Dil: en=97, tr=219
- Sınıf: IPM.Note=309, IPM.Schedule.Meeting.Request=3, IPM.Task=2, REPORT.IPM.Note.NDR=2
- Sorgu: 86 (ascii=8, attachment=8, crosslingual=10, date=11, exact=8, negative=8, partial=14, people=10, typo=9)

Her senaryo ailesinde bir hedef posta ve aynı konuda farklı ayrıntılı en az iki çeldirici vardır.
Yanıt zincirleri, iletilen postalar, imzalar, yasal uyarılar, uzun tutanaklar ve boş gövdeli postalar bulunur.
Görev ve teslim edilemedi (NDR) öğeleri aramada atlanmalıdır.

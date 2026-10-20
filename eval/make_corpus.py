"""Deterministic synthetic mail corpus and labeled queries for the HatırMail evaluation.

Run: python3 eval/make_corpus.py -> eval/corpus.json, eval/queries.json, eval/CORPUS.md (same bytes on every run).
All people, companies and addresses are fictional (reserved .example domains).
Each family = one target mail (first variant) + distractors on the same topic with other details.
"""
import json
import random
from datetime import datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
NOW = datetime(2026, 10, 14, 12, 0)
LO, HI = datetime(2024, 10, 1), datetime(2026, 10, 13, 23, 59)
rng = random.Random(20261014)
OWNER = ("Furkan Çakır", "furkan.cakir@atlas-ornek.example", "Satın Alma ve Lojistik Müdürü", "Atlas Örnek A.Ş.")
TR_MON = "Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık".split()
EN_MON = "January February March April May June July August September October November December".split()
TR_DAY = "Pazartesi Salı Çarşamba Perşembe Cuma Cumartesi Pazar".split()
EN_DAY = "Monday Tuesday Wednesday Thursday Friday Saturday Sunday".split()
FOLDERS = [{"id": "F-INBOX", "path": "\\Gelen Kutusu", "kind": "inbox"},
           {"id": "F-SENT", "path": "\\Gönderilmiş Öğeler", "kind": "sent"},
           {"id": "F-PROJ", "path": "\\Gelen Kutusu\\Projeler", "kind": "sub"},
           {"id": "F-ARCH", "path": "\\Arşiv", "kind": "other"}]
A = "Atlas Örnek A.Ş."
P = {
    "ahmet": ("Ahmet Demir", "ahmet.demir@hizli-lojistik.example", "Müşteri Temsilcisi", "Hızlı Lojistik A.Ş."),
    "zeynep": ("Zeynep Arslan", "zeynep.arslan@kuzey-kargo.example", "Operasyon Uzmanı", "Kuzey Kargo"),
    "mehmet": ("Mehmet Yılmaz", "mehmet.yilmaz@atlas-ornek.example", "Muhasebe Müdürü", A),
    "ayse": ("Ayşe Kaya", "ayse.kaya@atlas-ornek.example", "İnsan Kaynakları Uzmanı", A),
    "elif": ("Elif Şahin", "elif.sahin@atlas-ornek.example", "Satın Alma Uzmanı", A),
    "can": ("Can Öztürk", "can.ozturk@atlas-ornek.example", "BT Destek Uzmanı", A),
    "ipek": ("İpek Işık", "ipek.isik@atlas-ornek.example", "Proje Yöneticisi", A),
    "gokhan": ("Gökhan Çelik", "gokhan.celik@tedarikci.example", "Satış Müdürü", "Çelik Tedarik Ltd."),
    "selin": ("Selin Doğan", "selin.dogan@depo-cozum.example", "Satış Temsilcisi", "Depo Çözüm A.Ş."),
    "seda": ("Seda Güneş", "seda.gunes@atlas-ornek.example", "İSG Uzmanı", A),
    "murat": ("Murat Aydın", "murat.aydin@forklift-servis.example", "Servis Sorumlusu", "Forklift Servis Ltd."),
    "burak": ("Burak Koç", "burak.koc@atlas-ornek.example", "Satış Analisti", A),
    "hulya": ("Hülya Erdoğan", "hulya.erdogan@denetim-ornek.example", "Baş Denetçi", "Denetim Örnek Belgelendirme"),
    "oguz": ("Oğuz Kılıç", "oguz.kilic@atlas-ornek.example", "İdari İşler Sorumlusu", A),
    "kemal": ("Kemal Aksoy", "kemal.aksoy@atlas-ornek.example", "Finans Müdürü", A),
    "derya": ("Derya Polat", "derya.polat@atlas-ornek.example", "Ofis Yöneticisi", A),
    "nur": ("Nur Yıldız", "nur.yildiz@musteri-ornek.example", "Satın Alma Sorumlusu", "Yıldız Yapı Market"),
    "kerem": ("Kerem Uçar", "kerem.ucar@yapi-market.example", "Mağaza Müdürü", "Uçar Yapı"),
    "burcu": ("Burcu Tan", "burcu.tan@yazilim-dagitim.example", "Lisans Danışmanı", "Yazılım Dağıtım A.Ş."),
    "otel": ("Grand Ankara Otel Rezervasyon", "rezervasyon@grand-ankara.example", "", "Grand Ankara Otel"),
    "sarah": ("Sarah Miller", "sarah.miller@northwind-supply.example", "Account Manager", "Northwind Supply B.V."),
    "john": ("John Carter", "john.carter@northwind-supply.example", "Credit Controller", "Northwind Supply B.V."),
    "emma": ("Emma Schultz", "emma.schultz@bluepeak-software.example", "Customer Success Manager", "BluePeak Software GmbH"),
    "david": ("David Chen", "david.chen@bluepeak-software.example", "Security Engineer", "BluePeak Software GmbH"),
    "laura": ("Laura Rossi", "laura.rossi@eurotrans.example", "Pricing Manager", "EuroTrans S.p.A."),
    "michael": ("Michael Brown", "reservations@hotel-grand-berlin.example", "Reservations", "Hotel Grand Berlin"),
    "olivia": ("Olivia Green", "olivia.green@training-hub.example", "Course Coordinator", "Training Hub"),
    "tom": ("Tom Becker", "tom.becker@auditpro.example", "Senior Auditor", "AuditPro"),
}
DISC = {"tr": "Bu e-posta ve ekleri yalnızca alıcısına özel ve gizli bilgiler içerebilir. Bu mesajın muhatabı "
              "değilseniz lütfen göndereni bilgilendirip mesajı imha ediniz. {org} bu mesajın içeriğinden ve olası "
              "virüslerden sorumlu tutulamaz.",
        "en": "CONFIDENTIALITY NOTICE: This e-mail and any attachments are confidential and intended solely for the "
              "addressee. If you are not the intended recipient, please notify the sender and delete this message. "
              "Any unauthorized use or disclosure is prohibited. {org}"}
MOBILE = {"tr": "iPhone'umdan gönderildi", "en": "Sent from my iPhone"}
LONG = {"tr": [], "en": []}  # filled below
F = []  # families


def person(k):
    return OWNER if k == "owner" else P[k]


def fmt_dt(d, lang):
    if lang == "tr":
        return f"{d.day} {TR_MON[d.month - 1]} {d.year} {TR_DAY[d.weekday()]} {d:%H:%M}"
    return f"{EN_DAY[d.weekday()]}, {EN_MON[d.month - 1]} {d.day}, {d.year} {d:%I:%M %p}"


def signature(sender, lang):
    name, _, title, org = sender
    close = rng.choice(["Saygılarımla,", "İyi çalışmalar,", "Teşekkürler,"] if lang == "tr"
                       else ["Best regards,", "Kind regards,", "Thanks,"])
    phone = f"T: +90 212 555 {rng.randint(10, 99)} {rng.randint(10, 99)}" if org.endswith(("A.Ş.", "Ltd.", "Kargo", "Otel",
                                                                                         "Market", "Yapı", "Belgelendirme")) \
        else f"Phone: +31 20 {rng.randint(100, 999)} {rng.randint(1000, 9999)}"
    return "\n".join([close, name] + ([title] if title else []) + [org, phone])


def quoted(lang, orig, style):
    d = datetime.fromisoformat(orig["date"])
    to = "; ".join(n for n, _ in orig["to"])
    if style == "outlook" and lang == "tr":
        return (f"\n\nKimden: {orig['from_name']} <{orig['from_addr']}>\nGönderildi: {fmt_dt(d, 'tr')}\nKime: {to}\n"
                f"Konu: {orig['subject']}\n\n{orig['body']}")
    if style == "outlook":
        return (f"\n\nFrom: {orig['from_name']} <{orig['from_addr']}>\nSent: {fmt_dt(d, 'en')}\nTo: {to}\n"
                f"Subject: {orig['subject']}\n\n{orig['body']}")
    if style == "gmail":
        intro = (f"On {EN_DAY[d.weekday()][:3]}, {d.day} {EN_MON[d.month - 1][:3]} {d.year} at {d:%H:%M}, "
                 f"{orig['from_name']} <{orig['from_addr']}> wrote:" if lang == "en" else
                 f"{fmt_dt(d, 'tr')} tarihinde {orig['from_name']} <{orig['from_addr']}> şunu yazdı:")
        return "\n\n" + intro + "\n" + "\n".join("> " + x for x in orig["body"].split("\n"))
    return (f"\n\n-----Original Message-----\nFrom: {orig['from_name']}\nSent: {fmt_dt(d, 'en')}\nTo: {to}\n"
            f"Subject: {orig['subject']}\n\n{orig['body']}")


def forward_block(lang, m):
    d = datetime.fromisoformat(m["date"])
    to = "; ".join(f"{n} <{a}>" for n, a in m["to"])
    if lang == "tr":
        return (f"\n\n---------- İletilen ileti ----------\nKimden: {m['from_name']} <{m['from_addr']}>\n"
                f"Tarih: {fmt_dt(d, 'tr')}\nKonu: {m['subject']}\nKime: {to}\n\n{m['body']}")
    return (f"\n\n-----Original Message-----\nFrom: {m['from_name']} <{m['from_addr']}>\nSent: {fmt_dt(d, 'en')}\n"
            f"To: {to}\nSubject: {m['subject']}\n\n{m['body']}")


def long_body(lang, unique):
    paras = rng.sample(LONG[lang], 8)
    paras.insert(rng.randint(5, 7), unique)  # the distinctive detail sits in the second half
    head, tail = (("Merhaba ekip,\n\nBugünkü toplantının notlarını aşağıda paylaşıyorum.",
                   "Bir sonraki toplantı aynı gün ve saatte yapılacaktır. Eksik ya da hatalı bir nokta varsa lütfen yazın.")
                  if lang == "tr" else
                  ("Hi team,\n\nPlease find this week's status summary below.",
                   "The next update follows next week. Please send corrections before Friday."))
    return head + "\n\n" + "\n\n".join(f"{i + 1}. {p}" for i, p in enumerate(paras)) + "\n\n" + tail


MSGS = []


def add(folder, cls, d, sender, to, cc, subject, body, atts, lang):
    if folder == "F-INBOX" and d.year < 2026 or folder == "F-PROJ" and d < datetime(2025, 7, 1):
        folder = "F-ARCH"
    mod = min(d + timedelta(minutes=rng.randint(0, 600)), NOW)
    m = {"id": f"M{len(MSGS) + 1:04d}", "folder": folder, "cls": cls, "mod": mod.isoformat(timespec="seconds"),
         "subject": subject, "from_name": sender[0], "from_addr": sender[1], "to": [[t[0], t[1]] for t in to],
         "cc": [[c[0], c[1]] for c in cc], "date": d.isoformat(timespec="seconds"), "attachments": atts,
         "body": body, "lang": lang}
    MSGS.append(m)
    return m


def compose(lang, sender, new_text, var, tail=""):
    if var.get("empty"):
        return ""
    parts = [new_text]
    if var.get("sig", rng.random() < 0.5):
        parts.append(signature(sender, lang))
    if var.get("mob", rng.random() < 0.1):
        parts.append(MOBILE[lang])
    if var.get("disc", rng.random() < 0.2):
        parts.append(DISC[lang].format(org=sender[3]))
    return "\n\n".join(parts) + tail


def build_family(fam):
    lang, made = fam["lang"], []
    for var in fam["v"]:
        d = datetime.fromisoformat(var["d"])
        folder = var.get("f", fam.get("f", "F-INBOX"))
        other = person(var["p"])
        sender, to = (OWNER, [other]) if folder == "F-SENT" else (other, [OWNER])
        sv = {**fam.get("x", {}), **{k: x for k, x in var.items() if isinstance(x, str)}}
        sv.update(mon=TR_MON[d.month - 1], mon_en=EN_MON[d.month - 1], yr=str(d.year), ym=f"{d:%Y-%m}",
                  dd=str(d.day), due=(d + timedelta(days=15)).strftime("%d.%m.%Y"), name=sender[0])
        subject = var.get("s", fam["s"]).format(**sv)
        text = long_body(lang, sv["u"]) if fam.get("long") and "u" in sv else var.get("b", fam.get("b", "")).format(**sv)
        tail = ""
        fw = var["fw"] if "fw" in var else fam.get("fw")
        if fw:
            fs = person(fw["p"])
            tail = forward_block(lang, {"from_name": fs[0], "from_addr": fs[1], "date": fw["d"], "to": [list(sender[:2])],
                                        "subject": fw["s"].format(**sv), "body": fw["b"].format(**sv)})
        elif "re" in var:
            tail = quoted(lang, made[var["re"]], rng.choice(["outlook", "gmail"] if lang == "tr" else ["outlook", "gmail", "orig"]))
        body = compose(lang, sender, text, var, tail)
        atts = [a.format(**sv) for a in var.get("a", fam.get("a", []))]
        made.append(add(folder, var.get("cls", fam.get("cls", "IPM.Note")), d, sender, to,
                        [person(c) for c in var.get("cc", [])], subject, body, atts, lang))
    return made


# ---------------------------------------------------------------------------------------------------------------
F += [
    dict(k="kargo", lang="tr", s="{firm} kargo faturası – {mon} {yr}",
         b="Merhaba Furkan Bey,\n\n{mon} {yr} dönemine ait kargo faturamız ekte yer almaktadır. Toplam tutar {amt} TL, "
           "son ödeme tarihi {due}.\n\nİyi çalışmalar dileriz.",
         a=["Fatura_{ym}_{inv}.pdf"], x=dict(firm="Hızlı Lojistik"),
         v=[dict(p="ahmet", d="2026-09-12 09:31", amt="4.250", inv="HL4471", sig=1, disc=1),
            dict(p="ahmet", d="2026-03-11 10:02", amt="3.180", inv="HL3920"),
            dict(p="zeynep", d="2026-09-18 14:20", firm="Kuzey Kargo", amt="2.940", inv="KK118"),
            dict(p="mehmet", d="2026-09-02 10:00", s="FW: Hızlı Lojistik kargo faturası – Ağustos 2026",
                 b="Furkan, bu faturada fazla kesinti var gibi, kontrol eder misin?", amt="", inv="HL4102",
                 fw=dict(p="ahmet", d="2026-08-10T09:00:00", s="Hızlı Lojistik kargo faturası – Ağustos 2026",
                         b="Ağustos 2026 dönemine ait kargo faturamız ektedir. Toplam tutar 5.120 TL."))],
         q=[("date", "geçen ay Ahmet'ten gelen kargo faturası"), ("typo", "hızlı lojistk kargo faturası eylül")]),
    dict(k="forklift", lang="tr", s="Forklift tamir teklifi – {model}",
         b="Furkan Bey merhaba,\n\nDepodaki {model} forkliftin {part} arızası için tamir teklifimiz {amt} TL + KDV'dir. "
           "Parçalar {days} iş günü içinde temin edilir, işçilik fiyata dahildir.",
         a=["Teklif_{tno}.pdf"],
         v=[dict(p="murat", d="2026-07-08 11:15", model="Toyota 8FBE15", part="hidrolik pompa", amt="18.400", days="5",
                 tno="FS-2026-114", sig=1),
            dict(p="murat", d="2025-11-20 09:40", model="Linde E20", part="fren balatası", amt="6.250", days="3",
                 tno="FS-2025-388"),
            dict(p="murat", d="2026-09-03 16:05", model="Toyota 8FBE15", part="akü şarj ünitesi", amt="9.900", days="7",
                 tno="FS-2026-201")],
         q=[("partial", "forkliftin hidrolik arızası için gelen fiyat"),
            ("crosslingual", "forklift hydraulic pump repair quote")]),
    dict(k="izin", lang="tr", f="F-SENT", s="Yıllık izin talebi – {rng}",
         b="Merhaba Ayşe Hanım,\n\n{rng} tarihleri arasında {n} iş günü yıllık izin kullanmak istiyorum. İzinli olduğum "
           "sürede işlerime {backup} bakacak.\n\nBilgilerinize sunarım.",
         v=[dict(p="ayse", d="2026-07-01 10:20", rng="20–31 Temmuz", n="10", backup="Burak Koç", sig=1),
            dict(p="ayse", d="2025-12-15 09:05", rng="29 Aralık – 2 Ocak", n="4", backup="Seda Güneş"),
            dict(p="ayse", d="2026-04-02 14:10", rng="20–24 Nisan", n="5", backup="Oğuz Kılıç")],
         q=[("people", "Ayşe'ye gönderdiğim temmuz izin talebi"), ("ascii", "yillik izin talebi burak koc bakacak")]),
    dict(k="sunucu", lang="tr", s="Planlı sunucu bakımı – {when}",
         b="Merhaba,\n\n{when} gecesi saat 22:00 ile 02:00 arasında {system} sunucusunda planlı bakım yapılacaktır. "
           "Bu sürede {impact}. Lütfen açık işlemlerinizi önceden kaydedin.\n\nBT Ekibi",
         x=dict(system="ERP (SAP)", impact="ERP sistemine erişim olmayacaktır"),
         v=[dict(p="can", d="2026-10-07 17:30", when="10 Ekim Cumartesi", cc=["mehmet", "elif"]),
            dict(p="can", d="2026-05-14 16:00", when="16 Mayıs Cumartesi", system="e-posta",
                 impact="e-posta gönderimi gecikebilir"),
            dict(p="can", d="2025-10-09 15:45", when="11 Ekim Cumartesi"),
            dict(p="burak", d="2026-10-07 18:02", s="RE: Planlı sunucu bakımı – 10 Ekim Cumartesi",
                 b="Teşekkürler, bilgimiz var. Satış ekibine de iletiyorum.", re=0)],
         q=[("date", "geçen hafta gelen ERP bakım duyurusu"), ("typo", "sunucu bakımı erp sistemne erişim", (0, 2))]),
    dict(k="raf", lang="tr", s="Fiyat teklifi: {product}",
         b="Furkan Bey,\n\nTalep ettiğiniz {qty} adet {product} için birim fiyatımız {unit} TL, toplam {amt} TL'dir "
           "(KDV hariç). Teklifimiz {valid} gün geçerlidir; teslim süresi {days} gündür.",
         a=["Teklif_{tno}.pdf"], x=dict(valid="30", days="21"),
         v=[dict(p="gokhan", d="2026-03-04 10:12", product="endüstriyel raf sistemi", qty="40", unit="2.150",
                 amt="86.000", tno="TDR-0304", sig=1),
            dict(p="gokhan", d="2026-06-17 09:30", product="palet kaldırma arabası", qty="12", unit="7.800",
                 amt="93.600", tno="TDR-0617"),
            dict(p="selin", d="2026-03-19 13:45", product="endüstriyel raf sistemi", qty="40", unit="2.390",
                 amt="95.600", tno="DC-112", days="28")],
         q=[("date", "mart başında gelen raf teklifi"), ("people", "Gökhan'ın gönderdiği raf sistemi teklifi")]),
    dict(k="toplanti", lang="tr", f="F-PROJ", long=1, s="Atlas Depo Projesi – haftalık toplantı notları ({dd} {mon})",
         v=[dict(p="ipek", d="2026-09-24 18:10", u="Pilot uygulama için Gebze deposu seçildi; kurulum 2 Kasım'da "
                 "başlayacak ve ilk iki hafta yalnız gece vardiyasında çalışılacak."),
            dict(p="ipek", d="2026-09-17 18:05", u="Pilot uygulama için İzmir ve Gebze depoları karşılaştırıldı; "
                 "maliyet analizi tamamlanmadığı için karar bir sonraki toplantıya ertelendi."),
            dict(p="ipek", d="2026-10-01 18:20", u="El terminali tedarikçisi olarak Northwind seçildi; 40 cihazlık "
                 "sipariş Elif Şahin tarafından verilecek.")],
         q=[("partial", "toplantıda pilot için hangi depo seçilmişti"),
            ("exact", "Pilot uygulama için Gebze deposu seçildi")]),
]
F += [
    dict(k="otel", lang="tr", s="Rezervasyon onayı – {hotel}",
         b="Sayın Furkan Çakır,\n\n{hotel} için {stay} tarihleri arasında {nights} gecelik {room} rezervasyonunuz "
           "onaylanmıştır. Rezervasyon numaranız: {code}. Giriş saati 14:00, çıkış saati 12:00'dir.",
         x=dict(hotel="Grand Ankara Otel", room="tek kişilik oda"),
         v=[dict(p="otel", d="2026-06-02 15:40", stay="15–17 Haziran", nights="2", code="GA-55821"),
            dict(p="otel", d="2025-11-18 11:00", stay="24–25 Kasım", nights="1", code="GA-48810"),
            dict(p="ayse", d="2026-06-02 16:05", s="İLT: Rezervasyon onayı – Grand Ankara Otel",
                 b="Furkan Bey, Ankara fuarı için otel rezervasyonunuz yapıldı, onay aşağıdadır.",
                 fw=dict(p="otel", d="2026-06-02T15:40:00", s="Rezervasyon onayı – Grand Ankara Otel",
                         b="15–17 Haziran tarihleri arasında 2 gecelik tek kişilik oda rezervasyonunuz onaylanmıştır. "
                           "Rezervasyon numaranız: GA-55821."))],
         q=[("date", "yazın Ankara'daki otel rezervasyonu", (0, 2)),
            ("partial", "Ayşe bir otel rezervasyonu iletmişti galiba", 2)]),
    dict(k="butce", lang="tr", s="{yr} {q} bütçe taslağı",
         b="Merhaba,\n\n{yr} {q} dönemi bütçe taslağını ekte paylaşıyorum. Lojistik giderlerinde {pct} artış "
           "öngörülmektedir. Yorumlarınızı {dl} tarihine kadar iletmenizi rica ederim.",
         a=["Bütçe_{yr}_{q}.xlsx"],
         v=[dict(p="kemal", d="2026-06-10 09:00", q="Q3", pct="%12", dl="17 Haziran"),
            dict(p="kemal", d="2026-09-09 09:15", q="Q4", pct="%8", dl="16 Eylül", sig=1),
            dict(p="kemal", d="2025-12-08 10:30", q="Q1", pct="%15", dl="15 Aralık")],
         q=[("attachment", "bütçe excel dosyası olan mail, Q3"), ("date", "birkaç ay önce gelen bütçe taslağı")]),
    dict(k="sozlesme", lang="tr", s="Depo kiralama sözleşmesi – taslak v{ver}",
         b="Furkan Bey merhaba,\n\nDepo kiralama sözleşmesinin {ver}. taslağını ekte bulabilirsiniz. Bu sürümde "
           "{change}. Hukuk biriminin görüşü sonrası imzaya sunacağız.",
         a=["Sözleşme_Taslak_v{ver}.docx"],
         v=[dict(p="selin", d="2025-11-14 14:30", ver="3", change="kira artışı TÜFE ile sınırlandı ve fesih "
                 "bildirim süresi 90 güne çıkarıldı"),
            dict(p="selin", d="2025-10-22 11:20", ver="2", change="depo alanı 1.200 m² olarak düzeltildi"),
            dict(p="selin", d="2026-02-05 10:00", ver="4", change="sigorta bedeli tabloya eklendi")],
         q=[("date", "geçen yıl kasımda gelen sözleşme taslağı"), ("ascii", "sozlesme taslagi fesih bildirim suresi")]),
    dict(k="sifre", lang="tr", s="VPN şifre sıfırlama talebiniz",
         b="Merhaba Furkan,\n\nVPN hesabınız için şifre sıfırlama talebiniz alındı. Geçici şifreniz SMS ile "
           "gönderildi; ilk girişte değiştirmeniz gerekiyor. Talep no: {tno}.",
         v=[dict(p="can", d="2026-08-19 08:50", tno="BT-7782"),
            dict(p="can", d="2026-01-12 09:10", s="E-posta şifre sıfırlama talebiniz",
                 b="Merhaba Furkan,\n\nE-posta hesabınızın şifresi sıfırlandı. Yeni şifre en az 12 karakter olmalı. "
                   "Talep no: {tno}.", tno="BT-6120"),
            dict(p="can", d="2026-08-26 10:30", s="SAP kullanıcı kilidi açıldı",
                 b="Merhaba,\n\nSAP kullanıcınızın kilidi açıldı. Üç hatalı girişten sonra hesap tekrar kilitlenir. "
                   "Talep no: {tno}.", tno="BT-7810")],
         q=[("crosslingual", "VPN password reset request"), ("typo", "vpn şifre sıfırlma")]),
    dict(k="sikayet", lang="tr", s="Hasarlı teslimat hk. – Sipariş {order}",
         b="Merhaba,\n\n{order} numaralı siparişimizin {item} kısmı hasarlı teslim alındı. Fotoğrafları ekte "
           "gönderiyorum. Ürünlerin {days} gün içinde değiştirilmesini talep ediyoruz.",
         a=["hasar_{order}.jpg", "image001.png"],
         v=[dict(p="nur", d="2026-08-27 11:45", order="SP-20931", item="12 koli seramik karonun 3 kolisi", days="7"),
            dict(p="kerem", d="2026-05-06 15:10", order="SP-18874", item="paletli boya kutularının bir kısmı", days="10"),
            dict(p="nur", d="2026-02-17 09:25", order="SP-16602", item="2 adet lavabo bataryası", days="5"),
            dict(p="burak", d="2026-09-30 10:20", s="FW: Hasarlı teslimat hk. – Sipariş SP-21007", a=[],
                 b="Furkan Bey bilginize, müşteri yine yazmış.", order="SP-21007", item="", days="",
                 fw=dict(p="kerem", d="2026-09-30T09:40:00", s="Hasarlı teslimat hk. – Sipariş SP-21007",
                         b="SP-21007 numaralı siparişte 4 adet duş bataryası kırık geldi. Değişim bekliyoruz."))],
         q=[("partial", "müşteri kırık seramik koli şikayeti göndermişti"),
            ("crosslingual", "customer complaint about damaged ceramic tiles")]),
    dict(k="bordro", lang="tr", s="{mon} {yr} bordro ve ücret ödemesi",
         b="Değerli çalışma arkadaşlarımız,\n\n{mon} ayı maaş ödemeleri {pay} tarihinde hesaplarınıza yatırılacaktır. "
           "Bordrolarınızı İK portalından görüntüleyebilirsiniz. {extra}",
         x=dict(extra=""),
         v=[dict(p="ayse", d="2026-08-28 16:00", pay="31 Ağustos", extra="Bu ay yol yardımı tutarı 1.500 TL'ye "
                 "yükseltilmiştir."),
            dict(p="ayse", d="2026-07-29 16:10", pay="31 Temmuz"),
            dict(p="ayse", d="2026-09-28 16:05", pay="30 Eylül")],
         q=[("partial", "yol yardımı arttı diye bir İK duyurusu vardı")]),
    dict(k="egitim", lang="tr", s="{course} eğitimi – {when}",
         b="Merhaba,\n\n{course} eğitimi {when} günü saat {hour}'da {place} yapılacaktır. Katılım zorunludur, "
           "eğitim sonunda sertifika verilecektir.",
         x=dict(hour="09:30"),
         v=[dict(p="seda", d="2026-04-20 10:00", course="Forklift operatörü güvenlik", when="27 Nisan Pazartesi",
                 place="Gebze deposu toplantı salonunda", sig=1),
            dict(p="seda", d="2026-04-22 11:00", course="Yangın güvenliği", when="29 Nisan Çarşamba",
                 place="genel müdürlük konferans salonunda"),
            dict(p="seda", d="2025-10-13 09:30", course="İlk yardım", when="20 Ekim Pazartesi", place="çevrim içi")],
         q=[("ascii", "forklift operatoru guvenlik egitimi gebze")]),
    dict(k="tasima", lang="tr", s="Ofis taşınma planı – {phase}",
         b="Merhaba arkadaşlar,\n\nYeni ofise taşınma {when} yapılacaktır. {detail} Kişisel eşyalarınızı {box} "
           "tarihine kadar kolilere yerleştirmenizi rica ederiz.",
         v=[dict(p="oguz", d="2026-01-15 14:00", phase="1. aşama", when="24 Ocak Cumartesi",
                 detail="Maslak'taki yeni ofiste 5. kat bizim ekibe ayrıldı.", box="22 Ocak"),
            dict(p="oguz", d="2026-01-08 13:30", phase="duyuru", when="ay sonunda",
                 detail="Kat planı henüz kesinleşmedi.", box="20 Ocak"),
            dict(p="derya", d="2026-02-02 10:00", phase="2. aşama", when="7 Şubat Cumartesi",
                 detail="Arşiv dolapları bodrum kata taşınacak.", box="5 Şubat")],
         q=[("crosslingual", "office move plan which floor do we get")]),
    dict(k="lisans", lang="tr", s="Yazılım lisans yenileme – {product}",
         b="Furkan Bey merhaba,\n\n{product} lisanslarınızın süresi {exp} tarihinde doluyor. {n} kullanıcı için "
           "yenileme bedeli yıllık {amt} USD'dir.",
         v=[dict(p="burcu", d="2026-08-11 10:20", product="Adobe Acrobat Pro", exp="1 Eylül 2026", n="25", amt="4.875"),
            dict(p="burcu", d="2026-02-24 11:00", product="AutoCAD LT", exp="15 Mart 2026", n="5", amt="2.475"),
            dict(p="burcu", d="2025-08-12 09:50", product="Adobe Acrobat Pro", exp="1 Eylül 2025", n="20", amt="3.600")],
         q=[("people", "Burcu'nun gönderdiği 25 kullanıcılık Acrobat lisans yenileme")]),
    dict(k="denetim", lang="tr", s="ISO 9001 gözetim denetimi – {when}",
         b="Sayın yetkili,\n\n{when} tarihinde yapılacak ISO 9001 {kind} denetiminin planını ekte bulabilirsiniz. "
           "Denetim {days} gün sürecek; açılış toplantısı saat 09:00'dadır.",
         a=["Denetim_Planı_{yr}.pdf"],
         v=[dict(p="hulya", d="2026-05-04 15:20", when="18–19 Mayıs 2026", kind="gözetim", days="2"),
            dict(p="hulya", d="2025-05-12 14:00", when="26 Mayıs 2025", kind="gözetim", days="1"),
            dict(p="hulya", d="2026-09-14 10:10", s="ISO 9001 denetim raporu – {yr}",
                 b="Merhaba,\n\nMayıs ayında yapılan denetimin raporu ektedir. 2 minör uygunsuzluk tespit edilmiştir.",
                 a=["Denetim_Raporu_{yr}.pdf"])],
         q=[("attachment", "2026 ISO denetim planı pdf")]),
]
F += [
    dict(k="shipment", lang="en", s="Shipment delay – PO {po}",
         b="Hi Furkan,\n\nUnfortunately the shipment for PO {po} ({item}) is delayed at {place}. The new estimated "
           "arrival is {eta}. We apologise for the inconvenience.",
         v=[dict(p="sarah", d="2026-09-29 10:40", po="4500-7781", item="400 barcode scanners", place="Rotterdam port",
                 eta="8 October", sig=1),
            dict(p="sarah", d="2026-04-14 12:00", po="4500-6120", item="label printers", place="Hamburg",
                 eta="22 April"),
            dict(p="laura", d="2026-09-30 09:15", po="4500-7790", item="pallet wrap film", place="the Milan hub",
                 eta="6 October"),
            dict(p="elif", d="2026-08-24 11:05", s="Fwd: Customs hold – PO 4500-7802", b="FYI, see below.",
                 po="4500-7802", item="", place="", eta="",
                 fw=dict(p="sarah", d="2026-08-24T09:12:00", s="Customs hold – PO 4500-7802",
                         b="The customs office in Istanbul holds the 6 pallets of PO 4500-7802 until the missing "
                           "EUR.1 certificate arrives. Please send it today."))],
         q=[("people", "mail from Sarah about the delayed barcode scanner shipment"),
            ("crosslingual", "Rotterdam limanında geciken barkod okuyucu sevkiyatı"),
            ("typo", "shipmnet delay rotterdam"),
            ("partial", "gümrükte bekleyen paletler, EUR.1 belgesi eksikti", 3)]),
    dict(k="overdue", lang="en", s="Overdue invoice reminder – {inv}",
         b="Dear Furkan,\n\nOur records show that invoice {inv} for EUR {amt}, due on {due2}, is still unpaid. "
           "Please arrange payment or send us the remittance advice.",
         a=["{inv}.pdf"],
         v=[dict(p="john", d="2026-05-20 09:00", inv="INV-88213", amt="12,450.00", due2="30 April 2026"),
            dict(p="john", d="2025-11-03 10:20", inv="INV-80177", amt="3,980.00", due2="15 October 2025"),
            dict(p="john", d="2026-06-02 09:30", s="Payment received – {inv}",
                 b="Dear Furkan,\n\nThank you, we received the payment for invoice {inv}.", inv="INV-88213",
                 amt="12,450.00", due2="")],
         q=[("crosslingual", "Northwind'in 12.450 euroluk ödenmemiş fatura hatırlatması"), ("exact", "overdue invoice INV-88213")]),
    dict(k="renewal", lang="en", s="BluePeak WMS licence renewal – {yr}",
         b="Hello Furkan,\n\nYour BluePeak WMS subscription for {users} users renews on {ren}. The renewal price is "
           "EUR {amt} per year. Reply to confirm.",
         a=["BluePeak_Renewal_Quote_{yr}.pdf"],
         v=[dict(p="emma", d="2025-12-02 14:00", users="50", ren="1 January 2026", amt="18,000", sig=1),
            dict(p="emma", d="2024-12-03 13:30", users="40", ren="1 January 2025", amt="14,400"),
            dict(p="emma", d="2026-03-10 11:00", s="BluePeak WMS – 5 extra users",
                 b="Hello Furkan,\n\nAs requested, here is the quote for 5 additional users: EUR 1,800 per year.",
                 users="5", ren="", amt="1,800")],
         q=[("date", "licence renewal before new year"), ("attachment", "BluePeak renewal quote pdf for 50 users")]),
    dict(k="freight", lang="en", s="Freight rate offer {lane}",
         b="Dear Furkan,\n\nPlease find our rate for {lane}: EUR {amt} per full truck load, transit time {days} days, "
           "valid until {valid}.",
         v=[dict(p="laura", d="2026-02-16 09:45", lane="Istanbul – Milan", amt="2,850", days="5", valid="31 March"),
            dict(p="laura", d="2026-02-16 09:50", lane="Istanbul – Munich", amt="2,600", days="4", valid="31 March"),
            dict(p="laura", d="2025-09-01 10:00", lane="Istanbul – Milan", amt="2,640", days="6", valid="31 October")],
         q=[("crosslingual", "bu yıl gelen İstanbul Milano tır navlun fiyatı")]),
    dict(k="hotelen", lang="en", s="Booking confirmation – Hotel Grand Berlin",
         b="Dear Mr Çakır,\n\nThank you for your booking. Arrival {arr}, departure {dep}, {room}. Confirmation number "
           "{code}. Breakfast is included.",
         v=[dict(p="michael", d="2026-02-20 18:30", arr="10 March 2026", dep="13 March 2026", room="one superior room",
                 code="HGB-90417"),
            dict(p="michael", d="2025-02-11 17:00", arr="4 March 2025", dep="6 March 2025", room="one standard room",
                 code="HGB-71102"),
            dict(p="michael", d="2026-02-27 09:00", s="Pre-arrival information – Hotel Grand Berlin",
                 b="Dear Mr Çakır,\n\nWe look forward to welcoming you. Parking is available for EUR 25 per day.",
                 arr="", dep="", room="", code="")],
         q=[("crosslingual", "Berlin otel rezervasyon onayı, superior oda"), ("date", "booking confirmation last march Berlin")]),
    dict(k="security", lang="en", s="Security advisory – {topic}",
         b="Hello,\n\n{text}\n\nPlease contact support if you have questions.",
         v=[dict(p="david", d="2026-07-21 08:30", topic="mandatory MFA",
                 text="From 4 August, multi-factor authentication becomes mandatory for all admin accounts in the "
                      "BluePeak portal. Users without MFA will be locked out."),
            dict(p="david", d="2026-03-03 08:00", topic="TLS 1.0 shutdown",
                 text="On 31 March we disable TLS 1.0 and 1.1 on all BluePeak endpoints."),
            dict(p="emma", d="2026-07-22 12:00", s="RE: Security advisory – mandatory MFA",
                 b="Hi Furkan, yes, the MFA rule also applies to API service accounts.", re=0)],
         q=[("partial", "something about two factor login becoming required for admins")]),
    dict(k="trainingen", lang="en", s="Course registration – {course}",
         b="Dear Furkan,\n\nYou are registered for {course} on {when} ({fmt}). The joining link and the agenda are "
           "attached.",
         a=["Agenda_{code}.pdf"],
         v=[dict(p="olivia", d="2026-05-11 13:00", course="Advanced Excel for Logistics", when="2 June 2026",
                 fmt="online", code="AEL"),
            dict(p="olivia", d="2026-05-12 13:20", course="Lean Warehouse Management", when="16 June 2026",
                 fmt="Istanbul classroom", code="LWM"),
            dict(p="olivia", d="2025-10-06 10:00", course="Advanced Excel for Logistics", when="23 October 2025",
                 fmt="online", code="AEL")],
         q=[("crosslingual", "lojistik için ileri excel kurs kaydı haziran")]),
    dict(k="auditen", lang="en", s="Audit findings – {area}",
         b="Hi Furkan,\n\nOur visit found {n} findings in {area}. {main} Please send the corrective action plan by {by}.",
         v=[dict(p="tom", d="2026-08-05 16:40", area="inventory counting", n="3",
                 main="The most important: cycle count variances above 2% were not investigated.", by="4 September"),
            dict(p="tom", d="2026-08-05 16:45", area="supplier onboarding", n="1",
                 main="Two new suppliers had no signed code of conduct.", by="4 September"),
            dict(p="tom", d="2025-08-14 15:00", area="inventory counting", n="2",
                 main="Count sheets were not signed by the supervisor.", by="15 September")],
         q=[("partial", "the auditor said stock count differences were not checked"),
            ("typo", "audit finidngs inventory counting", (0, 2))]),
    dict(k="launch", lang="en", s="Product launch – {product}",
         b="Hi all,\n\nWe launch {product} on {when}. {detail} Marketing will share the visuals next week.",
         v=[dict(p="burak", d="2026-08-20 11:00", product="the Atlas SmartLocker", when="15 October",
                 detail="The first 200 units go to the Ankara stores."),
            dict(p="burak", d="2026-03-25 10:30", product="the Atlas Label Pro", when="10 April",
                 detail="Pre-orders open on 1 April."),
            dict(p="burak", d="2026-09-15 15:20", s="RE: Product launch – the Atlas SmartLocker",
                 b="Update: the launch moves to 22 October because of the packaging delay.", product="the Atlas SmartLocker",
                 when="15 October", detail="", re=0)],
         q=[("exact", "Atlas SmartLocker launch", (0, 2)), ("crosslingual", "akıllı dolap ürün lansmanı ertelendi mi", 2)]),
]
F += [
    dict(k="ise_alim", lang="tr", f="F-SENT", s="Mülakat daveti – {role}",
         b="Merhaba {cand},\n\n{role} pozisyonu için başvurunuz değerlendirildi. Sizi {when} saat {hour}'da "
           "{place} mülakata davet ediyoruz.",
         v=[dict(p="ayse", d="2026-09-21 10:00", cand="Deniz Bey", role="Depo Vardiya Amiri", when="28 Eylül",
                 hour="14:00", place="Gebze deposunda"),
            dict(p="ayse", d="2026-09-22 10:30", cand="Ece Hanım", role="Satın Alma Uzmanı", when="30 Eylül",
                 hour="11:00", place="genel müdürlükte"),
            dict(p="ayse", d="2026-03-09 09:00", cand="Tolga Bey", role="Depo Vardiya Amiri", when="16 Mart",
                 hour="10:00", place="Gebze deposunda")],
         q=[("date", "geçen ay depo vardiya amiri mülakat daveti")]),
    dict(k="arac", lang="tr", s="Araç bakım randevusu – {plate}",
         b="Merhaba,\n\n{plate} plakalı servis aracının {km} km bakımı için {when} günü randevu alındı. "
           "Araç {svc} servisine sabah 08:30'da teslim edilecek.",
         v=[dict(p="oguz", d="2026-10-13 09:20", plate="34 ABC 512", km="60.000", when="16 Ekim Cuma",
                 svc="Kartal"),
            dict(p="oguz", d="2026-06-08 10:00", plate="34 ABC 512", km="45.000", when="12 Haziran Cuma",
                 svc="Kartal"),
            dict(p="oguz", d="2026-10-02 14:00", plate="41 GBZ 208", km="90.000", when="9 Ekim Cuma",
                 svc="Gebze")],
         q=[("date", "dün gelen araç bakım randevusu")]),
    dict(k="avans", lang="tr", f="F-SENT", s="İş avansı kapama – {trip}",
         b="Mehmet Bey merhaba,\n\n{trip} seyahati için aldığım {adv} TL iş avansının kapama formu ve fişleri "
           "ekte. Harcanan tutar {spent} TL, iade edilecek tutar {ret} TL.",
         a=["Avans_Kapama_{code}.pdf"],
         v=[dict(p="mehmet", d="2026-06-22 11:30", trip="Ankara fuarı", adv="8.000", spent="6.740", ret="1.260",
                 code="ANK"),
            dict(p="mehmet", d="2026-03-17 10:00", trip="Berlin", adv="15.000", spent="14.300", ret="700", code="BER"),
            dict(p="mehmet", d="2025-11-27 09:30", trip="Ankara müşteri ziyareti", adv="5.000", spent="4.950",
                 ret="50", code="ANK2")],
         q=[("people", "Mehmet'e gönderdiğim Ankara fuarı avans kapama")]),
    dict(k="numune", lang="tr", s="Numune gönderimi – {item}",
         b="Merhaba Elif Hanım,\n\n{item} numuneleri bugün {cargo} ile gönderildi. Takip numarası {track}. "
           "Test sonuçlarını {days} gün içinde paylaşırsanız seri üretime geçebiliriz.",
         v=[dict(p="gokhan", d="2026-09-25 15:00", item="streç film (23 mikron)", cargo="Kuzey Kargo",
                 track="KK9918273", days="10"),
            dict(p="selin", d="2026-07-14 13:00", item="plastik palet", cargo="Hızlı Lojistik", track="HL5522019",
                 days="7"),
            dict(p="gokhan", d="2026-04-02 16:30", item="koli bandı", cargo="Kuzey Kargo", track="KK8711205",
                 days="5")],
         q=[("partial", "streç film numunesinin kargo takip numarası")]),
    dict(k="kalibrasyon", lang="tr", s="Kantar kalibrasyon sertifikası",
         b="Merhaba,\n\nGebze deposundaki {cap} ton kapasiteli kantarın kalibrasyonu {when} tarihinde yapıldı. "
           "Sertifika ektedir; bir sonraki kalibrasyon {next} tarihindedir.",
         a=["Kalibrasyon_Sertifikası_{yr}.pdf"],
         v=[dict(p="seda", d="2026-01-20 15:00", cap="60", when="19 Ocak 2026", next="Ocak 2027"),
            dict(p="seda", d="2025-01-23 14:10", cap="60", when="21 Ocak 2025", next="Ocak 2026"),
            dict(p="seda", d="2026-01-27 10:00", s="Terazi kalibrasyon planı",
                 b="Merhaba,\n\nDepodaki 12 hassas terazinin kalibrasyonu Şubat ayında yapılacaktır.",
                 cap="", when="", next="", a=[])],
         q=[("attachment", "kantar kalibrasyon sertifikası 2026")]),
]
# Families that only add noise (no target): newsletters, notices, calendar items, NDR.
NOISE_TR = [("Haftalık satış özeti – {n}. hafta", "Bu haftaki toplam satış {amt} TL oldu. Detaylar ekteki tabloda."),
            ("Yemekhane menüsü – {n}. hafta", "Bu haftanın menüsü ekte. Cuma günü tatlı ikramı vardır."),
            ("Otopark düzenlemesi", "Otoparktaki çizgi boyama çalışması nedeniyle {n} numaralı alanlar kapalıdır."),
            ("Toplantı odası rezervasyonu", "Toplantı odası {n} numaralı salon için rezervasyonunuz alındı."),
            ("Kırtasiye siparişi", "Kırtasiye siparişiniz {n} kalem olarak alındı, teslimat perşembe günü."),
            ("Doğum günü kutlaması", "Bugün {n}. kattaki arkadaşımızın doğum günü, saat 16:00'da pasta kesilecek."),
            ("Bilgi güvenliği hatırlatması", "Şüpheli e-postalardaki bağlantılara tıklamayın. Hatırlatma no: {n}."),
            ("Servis güzergahı değişikliği", "{n} numaralı servis güzergahında yol çalışması nedeniyle değişiklik var.")]
NOISE_EN = [("Weekly market update #{n}", "Freight rates on the main lanes stayed stable this week."),
            ("Webinar invitation: warehouse automation", "Join our free webinar on {n} November about automation."),
            ("Your monthly account statement", "Your statement for account {n} is available in the portal."),
            ("Newsletter – supply chain news", "Top stories this week: port congestion and fuel prices ({n}).")]

SYN_LONG_TR = ["Depo doluluk oranı geçen haftaya göre yüzde iki artarak yüzde seksen yedi seviyesine ulaştı.",
               "Sevkiyat planlamasında hafta sonu vardiyalarının yeniden düzenlenmesi konuşuldu.",
               "Barkod okuyucuların pil ömrü ile ilgili sahadan gelen şikayetler değerlendirildi.",
               "İş sağlığı ve güvenliği tarafında ayın kaza raporu sıfır olarak paylaşıldı.",
               "Müşteri iade süreçlerinin ortalama kapanma süresi dört günden üç güne indi.",
               "Yeni personelin oryantasyon programı için eğitim takvimi taslağı hazırlandı.",
               "Elektrik tüketimini azaltmak için depo aydınlatmasında sensör kullanımı önerildi.",
               "Tedarikçi performans karnesi çeyrek sonunda yönetimle paylaşılacak.",
               "Stok sayım farklarının nedenleri kategori bazında ayrıca incelenecek.",
               "Araç filosunun yakıt giderleri bütçeye göre yüzde beş üzerinde gerçekleşti.",
               "Raf etiketlerinin yeni standarda göre değiştirilmesi için iş planı yapıldı.",
               "Bir sonraki çeyrekte kapasite artışı için ek forklift kiralama seçeneği konuşuldu."]
SYN_LONG_EN = ["Warehouse occupancy rose two points to eighty-seven percent compared with last week.",
               "The team reviewed the weekend shift plan for outbound shipments.",
               "Field reports about the short battery life of the barcode scanners were discussed.",
               "The monthly health and safety report showed zero incidents.",
               "The average closing time of customer returns dropped from four days to three.",
               "A draft training calendar for the onboarding of new staff is ready.",
               "Motion sensors for the warehouse lighting were proposed to reduce power use.",
               "The supplier scorecard will be shared with management at quarter end.",
               "Stock count differences will be analysed per category.",
               "Fleet fuel costs were five percent above budget.",
               "A work plan to replace shelf labels with the new standard was agreed.",
               "Renting an extra forklift for next quarter's capacity was discussed."]
LONG["tr"], LONG["en"] = SYN_LONG_TR, SYN_LONG_EN
F += [
    dict(k="statusen", lang="en", f="F-PROJ", long=1, s="Atlas WMS rollout – weekly status ({mon_en} {dd})",
         v=[dict(p="emma", d="2026-09-11 17:00", u="Decision: the go-live for the Izmir warehouse moves to "
                 "9 November because the label printer drivers failed the acceptance test."),
            dict(p="emma", d="2026-09-04 17:10", u="The Izmir warehouse acceptance test starts next week; all label "
                 "printer drivers are installed."),
            dict(p="emma", d="2026-09-18 17:05", u="The Gebze go-live went live as planned with 12 users; no "
                 "critical defects were reported.")],
         q=[("partial", "why was the Izmir go-live postponed")]),
    dict(k="kalite", lang="tr", f="F-PROJ", long=1, s="Kalite komitesi toplantı tutanağı – {mon} {yr}",
         v=[dict(p="seda", d="2026-06-30 17:30", u="Müşteri şikayetlerinde en sık neden yanlış etiketleme çıktı; "
                 "Temmuz'dan itibaren çift kontrol uygulamasına geçilecek."),
            dict(p="seda", d="2026-05-29 17:20", u="Şikayet kayıtlarının yeni CRM sistemine aktarımı tamamlandı."),
            dict(p="seda", d="2026-07-31 17:15", u="Çift kontrol uygulaması ilk ayında hatalı etiketlemeyi yarı yarıya "
                 "azalttı.")],
         q=[("ascii", "kalite komitesi yanlis etiketleme cift kontrol")]),
]
QUERIES_EXTRA = [  # (category, query, family key, variant index) for queries that need other targets
    ("exact", "Yıllık izin talebi 20–31 Temmuz", "izin", 0),
    ("exact", "Hasarlı teslimat SP-20931", "sikayet", 0),
    ("exact", "Course registration Advanced Excel for Logistics June", "trainingen", 0),
    ("exact", "Grand Ankara Otel rezervasyon onayı GA-55821", "otel", (0, 2)),
    ("exact", "Freight rate offer Istanbul – Milan 2,850", "freight", 0),
    ("ascii", "vpn sifre sifirlama talebi", "sifre", 0),
    ("ascii", "hasarli teslimat seramik koli", "sikayet", 0),
    ("ascii", "ofis tasinma plani 5. kat maslak", "tasima", 0),
    ("ascii", "egitim forklift guvenlik sertifika nisan", "egitim", 0),
    ("typo", "bütçe taslğı Q3 lojistik", "butce", 0),
    ("typo", "rezervasyn onayı grand ankara haziran", "otel", (0, 2)),
    ("typo", "yol yardimi bordrro ağustos", "bordro", 0),
    ("typo", "lisans yenilme acrobat 25 kullanıcı", "lisans", 0),
    ("partial", "geçen ay gelen MFA zorunluluğu duyurusu", "security", 0),
    ("date", "2 hafta önce gelen tarayıcı sevkiyatı gecikmesi", "shipment", 0),
    ("partial", "araç bakımı 60 bin km", "arac", 0),
    ("partial", "numune kargo takip numarası Gökhan", "numune", 0),
    ("partial", "denetçi stok sayım farkı demişti", "auditen", 0),
    ("people", "Murat'tan gelen forklift teklifi Toyota hidrolik", "forklift", 0),
    ("people", "John's payment reminder for invoice 88213", "overdue", 0),
    ("people", "Laura'nın şubatta gönderdiği Milano navlun fiyatı", "freight", 0),
    ("people", "Emma'nın gönderdiği İzmir canlıya geçiş ertelemesi", "statusen", 0),
    ("people", "Can'dan gelen VPN şifre maili", "sifre", 0),
    ("attachment", "Fatura_2026-09 pdf Hızlı Lojistik", "kargo", 0),
    ("attachment", "Sözleşme_Taslak_v3 docx", "sozlesme", 0),
    ("attachment", "avans kapama formu ANK pdf", "avans", 0),
    ("attachment", "Agenda AEL pdf course", "trainingen", (0, 2)),
    ("negative", "uzay mekiği fırlatma takvimi", None, None),
    ("negative", "düğün davetiyesi ve nişan fotoğrafları", None, None),
    ("negative", "kripto para cüzdanı şifresi", None, None),
    ("negative", "the quarterly dividend payout to shareholders", None, None),
    ("negative", "veteriner aşı randevusu kedi", None, None),
    ("negative", "swimming pool membership renewal", None, None),
    ("negative", "forklift lastiği satın alma ihalesi Antalya", None, None),
    ("negative", "otel havuzu bakım ihalesi", None, None),
]



def rand_date(lo=LO, hi=HI):
    span = int((hi - lo).total_seconds() // 60)
    d = lo + timedelta(minutes=rng.randint(0, span))
    return d.replace(hour=rng.randint(8, 18), second=0, microsecond=0)


def build():
    fams = {}
    for fam in F:
        fams[fam["k"]] = build_family(fam)
    owner_ccs = [p for k, p in P.items() if p[1].endswith("atlas-ornek.example")]
    for i in range(100):  # Turkish noise
        subj, body = NOISE_TR[i % len(NOISE_TR)]
        v = {"n": str(rng.randint(2, 48)), "amt": f"{rng.randint(120, 990)}.{rng.randint(100, 999)}"}
        sender = rng.choice(owner_ccs)
        folder = rng.choice(["F-INBOX", "F-PROJ", "F-PROJ", "F-ARCH"])
        text = "Merhaba,\n\n" + body.format(**v)
        add(folder, "IPM.Note", rand_date(), sender, [OWNER], [], subj.format(**v),
            compose("tr", sender, text, {}), ["image001.png"] if rng.random() < 0.2 else [], "tr")
    for i in range(50):  # English noise
        subj, body = NOISE_EN[i % len(NOISE_EN)]
        v = {"n": str(rng.randint(2, 48))}
        sender = rng.choice([P["sarah"], P["laura"], P["olivia"], P["emma"]])
        add(rng.choice(["F-INBOX", "F-ARCH"]), "IPM.Note", rand_date(), sender, [OWNER], [], subj.format(**v),
            compose("en", sender, "Hello,\n\n" + body.format(**v), {}), [], "en")
    for i in range(56):  # sent noise
        other = rng.choice(list(P.values()))
        lang = "en" if other[1].split("@")[1] in ("northwind-supply.example", "bluepeak-software.example",
                                                  "eurotrans.example", "hotel-grand-berlin.example",
                                                  "training-hub.example", "auditpro.example") else "tr"
        text = rng.choice(["Teşekkürler, aldım. Kontrol edip dönüş yapacağım.", "Toplantı için uygun saat 14:00 olabilir mi?",
                           "Ekteki dosyayı inceleyip görüşlerinizi paylaşır mısınız?", "Bilginize, konu takipteyiz."]
                          if lang == "tr" else
                          ["Thanks, received. I will check and come back to you.", "Could we meet at 2 pm on Thursday?",
                           "Please review the attached file and share your comments.", "Noted, we follow up on this."])
        add("F-SENT", "IPM.Note", rand_date(), OWNER, [other], [], rng.choice(["Bilgi", "Takip", "Görüşme", "Dönüş"]
            if lang == "tr" else ["Follow-up", "Quick question", "Meeting", "Update"]), compose(lang, OWNER, text, {}), [], lang)
    # Meeting requests (searchable), tasks and NDR (skipped by the search).
    for subj, d, lang in [("Haftalık satın alma koordinasyon toplantısı", "2026-09-07 09:00", "tr"),
                          ("Quarterly business review with Northwind", "2026-07-02 15:00", "en"),
                          ("Depo güvenlik turu", "2026-05-18 10:00", "tr")]:
        add("F-INBOX", "IPM.Schedule.Meeting.Request", datetime.fromisoformat(d), P["ipek"], [OWNER], [], subj,
            "Toplantı daveti. Yer: Toplantı odası 3." if lang == "tr" else "Meeting invitation. Room: Board room.", [], lang)
    for subj in ("Görev: tedarikçi listesini güncelle", "Görev: sayım planını hazırla"):
        add("F-INBOX", "IPM.Task", rand_date(datetime(2026, 1, 1)), OWNER, [OWNER], [], subj, "Görev ayrıntısı.", [], "tr")
    for d in ("2026-04-03 10:00", "2025-12-12 16:30"):
        add("F-INBOX", "REPORT.IPM.Note.NDR", datetime.fromisoformat(d), ("Mail Delivery System",
            "postmaster@atlas-ornek.example", "", A), [OWNER], [], "Teslim edilemedi: Sipariş onayı",
            "İleti şu alıcılara teslim edilemedi: siparis@kapanmis-firma.example", [], "tr")
    for subj, atts in [("Tarama", ["Scan_20260611_0931.pdf"]), ("", ["Fatura_2026-06_TR8812.pdf", "image001.png"]),
                       ("Scanned document", ["scan_0007.pdf"])]:  # empty-body mails
        add("F-INBOX", "IPM.Note", rand_date(datetime(2026, 1, 1)), P["derya"], [OWNER], [], subj, "", atts, "tr")
    return fams


def collect_queries(fams):
    qs = []
    for fam in F:
        for q in fam.get("q", []):
            cat, text = q[0], q[1]
            idx = q[2] if len(q) > 2 else 0
            idxs = idx if isinstance(idx, tuple) else (idx,)
            qs.append({"q": text, "lang": "en" if fam["lang"] == "en" and cat != "crosslingual"
                       or fam["lang"] == "tr" and cat == "crosslingual" else "tr",
                       "category": cat, "targets": [fams[fam["k"]][i]["id"] for i in idxs], "family": fam["k"]})
    for cat, text, key, i in QUERIES_EXTRA:
        qs.append({"q": text, "lang": "en" if all(ord(c) < 128 for c in text) and cat != "ascii" else "tr",
                   "category": cat, "family": key or "",
                   "targets": [] if key is None else [fams[key][j]["id"] for j in (i if isinstance(i, tuple) else (i,))]})
    for n, q in enumerate(qs, 1):
        q["id"] = f"Q{n:03d}"
    return [{k: q[k] for k in ("id", "q", "lang", "category", "targets", "family")} for q in qs]


def validate(msgs, qs):
    ids = [m["id"] for m in msgs]
    assert len(ids) == len(set(ids))
    assert 300 <= len(msgs) <= 420, len(msgs)
    fids = {f["id"] for f in FOLDERS}
    by = {m["id"]: m for m in msgs}
    for m in msgs:
        assert m["folder"] in fids
        assert LO <= datetime.fromisoformat(m["date"]) <= HI, m["date"]
        assert m["mod"] >= m["date"]
        assert ".example" in m["from_addr"]
    for f in fids:
        assert sum(m["folder"] == f for m in msgs) >= 40, (f, sum(m["folder"] == f for m in msgs))
    assert len({q["id"] for q in qs}) == len(qs)
    mins = {"exact": 8, "ascii": 8, "typo": 8, "date": 8, "partial": 10, "crosslingual": 8, "people": 8,
            "attachment": 6, "negative": 8}
    for c, n in mins.items():
        assert sum(q["category"] == c for q in qs) >= n, (c, sum(q["category"] == c for q in qs))
    for q in qs:
        assert q["category"] in mins, q
        for t in q["targets"]:
            assert by[t]["cls"] in ("IPM.Note", "IPM.Schedule.Meeting.Request"), q
        assert q["category"] == "negative" or q["targets"], q


def main():
    fams = build()
    MSGS.sort(key=lambda m: m["date"])
    old_to_new = {m["id"]: f"M{i:04d}" for i, m in enumerate(MSGS, 1)}
    for m in MSGS:
        m["id"] = old_to_new[m["id"]]
    qs = collect_queries(fams)
    validate(MSGS, qs)
    corpus = {"now": NOW.isoformat(timespec="seconds"), "store": {"id": "STORE-SYNTH-1", "name": "Sentetik Posta (Atlas Örnek)"},
              "folders": FOLDERS, "messages": MSGS}
    (HERE / "corpus.json").write_text(json.dumps(corpus, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    (HERE / "queries.json").write_text(json.dumps(qs, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    from collections import Counter
    fc, lc, cc = Counter(m["folder"] for m in MSGS), Counter(m["lang"] for m in MSGS), Counter(m["cls"] for m in MSGS)
    qc = Counter(q["category"] for q in qs)
    lines = ["# Sentetik test verisi", "", "Bu veri tamamen kurgusaldır. Kişiler, firmalar ve adresler gerçek değildir "
             "(`.example` alan adı). Yeniden üretmek için: `python3 eval/make_corpus.py`.", "",
             f"- Toplam posta: {len(MSGS)}", "- Klasörler: " + ", ".join(f"{k}={v}" for k, v in sorted(fc.items())),
             "- Dil: " + ", ".join(f"{k}={v}" for k, v in sorted(lc.items())),
             "- Sınıf: " + ", ".join(f"{k}={v}" for k, v in sorted(cc.items())),
             f"- Sorgu: {len(qs)} (" + ", ".join(f"{k}={v}" for k, v in sorted(qc.items())) + ")", "",
             "Her senaryo ailesinde bir hedef posta ve aynı konuda farklı ayrıntılı en az iki çeldirici vardır.",
             "Yanıt zincirleri, iletilen postalar, imzalar, yasal uyarılar, uzun tutanaklar ve boş gövdeli postalar bulunur.",
             "Görev ve teslim edilemedi (NDR) öğeleri aramada atlanmalıdır."]
    (HERE / "CORPUS.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"messages={len(MSGS)} folders={dict(fc)} lang={dict(lc)} cls={dict(cc)}")
    print(f"queries={len(qs)} {dict(qc)}")


if __name__ == "__main__":
    main()

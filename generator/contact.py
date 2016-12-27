from model.contact import Contact
import random
import string
import os.path #для работы с путями до файлов
import jsonpickle
import getopt # для использования опций из командной строки
import sys #для получения доступа к этим опциям


# описание получение опций из оф. документации питона
try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of contacts", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

# указываем дефольные значения
n = 2
f = "data/contacts.json"

for o, a in opts:
    if o == "-n": #если название опции равно -n
        n = int(a) #значит в ней задается количество групп
    elif o == "-f":
        f = a



#случайные строки для имени
def random_string(prefix, maxlen):
    symbols = string.ascii_letters + "-"*10
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])

#случайные строки для имейла
def random_string_email(prefix, maxlen):
    symbols = string.ascii_letters + string.digits
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))]) + "@" + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))]) + "." + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])

#случайные строки для телефона
def random_string_phone(prefix, maxlen):
    symbols = string.digits
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])

#случайные строки для адреса
def random_string_add(prefix, maxlen):
    symbols = string.ascii_letters + string.digits + "-"*10
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])


testdata = [Contact(F_name=random_string("fn_", 5), L_name=random_string("ln_", 5),
                    C_address=random_string_add("addr", 5), H_phone=random_string_phone("hom_", 7),
                    M_phone=random_string_phone("mob_", 10), W_phone=random_string_phone("wor_", 8),
                    S_phone=random_string_phone("sec_", 9), C_email=random_string_email("email_", 4))
            for i in range(n)
            ]

# сохраняем сгенерированные данные в фаил
file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)  # путь к файлу

# открываем фаил
with open(file, "w") as out:
    # dumps - превращает данные в строку вида json
    #out.write(json.dumps(testdata, default=lambda x: x.__dict__, indent=2)) #т.к. json не знает как преобразовывать объъекты типа group в словарь, мы ему подскажем - лямбда функцией. он сначала выполнит ее, а потом будет преобразовывать
    jsonpickle.set_encoder_options("json", indent=2)
    out.write(jsonpickle.encode(testdata))

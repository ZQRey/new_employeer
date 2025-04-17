import ssl
from ldap3 import Server, Connection, ALL, MODIFY_REPLACE, SUBTREE, Tls

# Подключение
AD_SERVER   = 'gp1.loc'
AD_PORT     = 389
AD_USER     = 'service_user@gp1.loc'
AD_PASSWORD = 'Zx213141@'

def process_employee_data(raw):
    # Словарь из формы
    fist_name = raw['first_name'].strip()
    last_name = raw['last_name'].strip()
    middle_name = raw['middle_name'].strip()
    login = (last_name + "_" + fist_name[0]).lower()
    return {
        'iin': raw['iin'].strip(),
        'first_name': fist_name,
        'last_name': last_name,
        'middle_name': middle_name,
        'full_name': f"{last_name} {fist_name} {middle_name}",
        'login': login,
        'password': 'Aa0000',
        'specialty': raw['specialty'].strip(),
        'phone': raw['phone'].strip(),
        'email': raw['email'].strip(),
    }

def create_ad_user(data):
    tls = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLSv1_2)
    server = Server(AD_SERVER, port=AD_PORT, use_ssl=False, get_info=ALL, tls=tls)
    try:
        conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
        if not conn.start_tls():
            conn.unbind()
            return False
    except:
        return False

    dn = f"CN={data['full_name']},{data['department_dn']}"
    attrs = {
        'objectClass':['top','person','organizationalPerson','user'],
        'givenName':data['first_name'],
        'sn':data['last_name'],
        'displayName':data['full_name'],
        'sAMAccountName':data['login'],
        'userPrincipalName':f"{data['login']}@gp1.local",
        'mail':data['email']
    }
    if not conn.add(dn, attributes=attrs):
        conn.unbind()
        return False

    if not conn.extend.microsoft.modify_password(dn, data['password'], None):
        conn.unbind()
        return False

    conn.modify(dn, {'pwdLastSet':[(MODIFY_REPLACE,[0])]})
    conn.modify(dn, {'userAccountControl':[(MODIFY_REPLACE,[512])]})
    conn.unbind()
    return True

def get_ad_departments():
    allowed = {
      "adm":"Административно-управленческая часть",
      "ago":"Акушерско - гинекологическое отделение (поликлиника)",
      "apt":"аптека (поликлиника)",
      "buh":"Бухгалтерия",
      "csz1":"Центр семейного здоровья (поликлиника) ЦСЗ 1",
      "csz2":"Центр семейного здоровья (поликлиника) ЦСЗ 2",
      "it":"ИТ отдел",
      "kdo":"Консультативно - диагностическое отделение (поликлиника)",
      "lab":"Лаборатория",
      "osh":"Школьное отделение (поликлиника)",
      "otdelkadrov":"Отдел кадров",
      "procedur":"Отделение профилактики (поликлиника)",
      "prof":"Консультативное отделение (поликлиника)",
      "reg":"Регистратура (поликлиника)",
      "skoray":"Скорая помощь",
      "temp":"Временные логины"
    }
    server = Server(AD_SERVER, port=AD_PORT, use_ssl=False, get_info=ALL)
    conn = Connection(server, user=AD_USER, password=AD_PASSWORD, auto_bind=True)
    conn.search('DC=gp1,DC=loc',
                '(&(objectClass=organizationalUnit))',
                SUBTREE,
                attributes=['distinguishedName'])
    deps = []
    for e in conn.entries:
        dn = e.distinguishedName.value
        key = dn.split(",")[0].replace("OU=","").strip().lower()
        if key in allowed:
            deps.append((dn, allowed[key]))
    conn.unbind()
    deps.sort(key=lambda x: x[1])
    return deps

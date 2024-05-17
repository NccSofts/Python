from datetime import datetime
from datetime import timedelta
from dateutil import parser
import json
import requests
import sqlalchemy
import traceback

import biblioteca
import parametros

# pip install mysql-connector-client

global_tray_url = parametros.global_tray_url
global_tray_consumer_key = parametros.global_tray_consumer_key
global_tray_consumer_secret = parametros.global_tray_consumer_secret
global_tray_code = parametros.global_tray_code
global_tray_access_token = ""
global_tray_refresh_token = ""

global_mysql = parametros.global_mysql

global_headers = { \
    "Accept": "application/json; charset=utf-8", \
}

global_sql = True
#global_sql = False

#def main():
#    ...

def authenticate():
    global global_tray_access_token
    global global_tray_refresh_token

    #print("\nauthenticate()")
    url = global_tray_url + "/auth"
    data = { \
        "consumer_key": global_tray_consumer_key, \
        "consumer_secret": global_tray_consumer_secret, \
        "code": global_tray_code \
    }
    response = requests.post(url=url, headers=global_headers, json=data)
    #print_json(response.json())
    global_tray_access_token = response.json().get("access_token")
    global_tray_refresh_token = response.json().get("refresh_token")

def refresh():
    global global_tray_access_token
    global global_tray_refresh_token

    #print("\nrefresh()")
    url = global_tray_url + f"/auth?refresh_token={global_tray_refresh_token}"
    response = requests.get(url=url, headers=global_headers)
    #print_json(response.json())
    global_tray_access_token = response.json().get("access_token")
    global_tray_refresh_token = response.json().get("refresh_token")

def order_json(id, cupons):
    url = global_tray_url + "/orders" \
        + f"/{id}/complete" \
        + f"?access_token={global_tray_access_token}" 
    response = requests.get(url=url, headers=global_headers)

    order = response.json().get("Order")
    if order is None:
        print("?")
        return

    ok = False
    coupon = order.get("coupon")
    coupon_code = ""
    if coupon is not None:
        regis = coupon.get("code").upper()
        #if regis.startswith("REGISTADEU") \
        #    or regis.startswith("BLACKREGIS") \
        #    or regis.startswith("REGISHENDRIX") \
        #    or regis.startswith("REGISKISS") \
        #    or regis.startswith("REGISLEE20") \
        #    or regis.startswith("REGISLEE30") \
        #    or regis.startswith("REGISRUSH"):
        if regis in cupons:
            coupon_code = coupon.get("code")
            ok = True

    created = order.get("date")
    modified = order.get("modified")
    status = order.get("status")
    print(f"{id} {created} {modified} {status} [{coupon_code}]")
    if ok:
        sql = sqlalchemy.text("SELECT COUNT(*) FROM Tray_Order_JSON WHERE id = :id")
        result = connection.execute(sql, { "id": id })
        if result.scalar() == 0:
            print("INSERT")
            sql = sqlalchemy.text("""
INSERT INTO Tray_Order_JSON
(
    id,
    created,
    modified,
    status,
    document
)
VALUES
(
    :id,
    :created,
    :modified,
    :status,
    :document
)
""")

            #"created": biblioteca.datetime_to_isoformat(created), \
            #"modified": biblioteca.datetime_to_isoformat(modified), \

            insert = { \
                "id": id, \
                "created": created, \
                "modified": modified, \
                "status": status, \
                "document": json.dumps(order) \
            }
            connection.execute(sql, insert)
            connection.commit()
        else: # status
            print("UPDATE")
            sql = sqlalchemy.text("""
UPDATE Tray_Order_JSON
SET status = :status
WHERE id = :id
""")
            update = { \
                "id": id, \
                "status": status \
            }
            connection.execute(sql, update)
            connection.commit()

def orders_json(connection, cupons):
    sql = sqlalchemy.text("SELECT MAX(modified) FROM Tray_Order_JSON LIMIT 1")
    result = connection.execute(sql)

    start_modified = result.scalar()
    if start_modified is None:
        start_modified = parser.parse("2023-01-01T00:00:00.000000Z")
    else: # status
        start_modified = start_modified - timedelta(days=7)

    page = 1
    #limit = 10
    limit = 50
    while page is not None:
        print(f"página {page}")

        modified = start_modified.strftime("%Y-%m-%d")
        url = global_tray_url + "/orders" \
            + f"?access_token={global_tray_access_token}" \
            + f"&limit={limit}" \
            + f"&modified={modified},2099-12-31" \
            + f"&page={page}" \
            + "&sort=id_asc"

        #    + f"&modified={modified},2099-12-31" \


        response = requests.get(url=url, headers=global_headers)
        data = response.json()
        #"paging": {
        #    "limit": 2,
        #    "maxLimit": 50,
        #    "offset": 0,
        #    "page": 1,
        #    "total": 12
        #}
        #
        #{
        #    "limit": 23,
        #    "maxLimit": 50,
        #    "offset": 322, offset > total
        #    "page": 15,
        #    "total": 301
        #}
        if "paging" in data:
            paging = data.get("paging")
            #print_json(paging)
            if "total" in paging:
                offset = paging.get("offset")
                total = paging.get("total")
                if offset > total:
                    page = None
                else:
                    page += 1
            else:
                page = None
        else:
            page = None

        if page is not None and "Orders" in data:
            orders = data.get("Orders")
            for o in orders:
                order = o.get("Order")
                id = order.get("id")
                print(id)
                order_json(id, cupons)

def print_json(data):
    print(json.dumps(data, sort_keys=True, indent=4))

try:

    #if __name__ == '__main__':
    #    main()

    # -> main()
    authenticate()

    engine = sqlalchemy.create_engine(global_mysql)

    print("Tray_Order_JSON")
    with engine.connect() as connection:
        sql = sqlalchemy.text("SELECT COALESCE(GROUP_CONCAT(Codigo SEPARATOR '|'), '') Cupons FROM Tray_Cupom ORDER BY Codigo")
        result = connection.execute(sql)
        cupons = result.scalar()
        if cupons is not None:
            print(cupons)
            orders_json(connection, cupons)

    print("Tray_Order | Tray_Ordem_Item")
    raw_connection = engine.raw_connection()
    try:
        c = raw_connection.cursor() # BEGIN TRANSACTION
        c.callproc("SP_Tray_Order")    
        c.close()
        raw_connection.commit() # COMMIT TRANSACTION
    finally:
        raw_connection.close()
    # -> main()

    biblioteca.ok("tray-read.py")
except Exception as exception:
    print(exception)
    traceback.print_exc()
    biblioteca.excecao(exception, "tray-read.py")

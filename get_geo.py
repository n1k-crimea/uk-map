import requests
from cfg import API_KEY_YA_MAP


def get_coord(addres):
    a = requests.get(
        'https://geocode-maps.yandex.ru/1.x/?format=json&apikey={}&geocode={}'.format(API_KEY_YA_MAP, addres))
    d = dict(a.json())
    try:
        point = d['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos'].split(' ')
    except IndexError:
        return [None, None]
    coord = list(map(float, point))
    coord.reverse()
    print(addres, coord)
    return coord

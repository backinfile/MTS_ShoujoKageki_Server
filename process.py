import os.path

from pandas import DataFrame
import json
import openpyxl

card_data_map = {}  # card:{ascension:{viewCnt,pickCnt,winCnt,pickFloorSum}}
combat_data_map = {}  # ascension:{victory:cnt, lose:cnt, loseLayerSum:cnt, perFloor:{}}

card_name_map = {}


def init_card_name_map():
    with open('ShoujoKageki-Card-Strings.json', encoding='utf-8') as f:
        content = json.load(f)
        for card, strings in content.items():
            card_name_map[card] = strings['NAME']


def processJson():
    files = os.listdir('data')
    for file_name in files:
        if not file_name.endswith('.json'):
            continue
        with open(os.path.join('data', file_name), 'r') as f:
            content = json.load(f)
            floor_reached = content['event']['floor_reached']
            mods = content['event']['mods']
            if floor_reached < 3:
                continue
            card_choices = content['event']['card_choices']
            ascension_level = int(content['event']['ascension_level'])
            victory = content['event']['victory']

            set_combat_data_default(ascension_level)
            if victory:
                combat_data_map[ascension_level]['victory'] += 1
            else:
                combat_data_map[ascension_level]['lose'] += 1
                combat_data_map[ascension_level]['loseLayerSum'] += floor_reached
                combat_data_map[ascension_level]['perFloor'][floor_reached] += 1

            for choice in card_choices:
                if int(choice['floor']) <= 0:
                    continue
                for card in choice['not_picked']:
                    card = get_raw_card_name(card)
                    set_card_data_default(card, ascension_level)
                    card_data_map[card][ascension_level]['viewCnt'] += 1
                card = choice['picked']
                card = get_raw_card_name(card)
                if card:
                    set_card_data_default( card, ascension_level )
                    card_data_map[card][ascension_level]['viewCnt'] += 1
                    card_data_map[card][ascension_level]['pickCnt'] += 1
                    card_data_map[card][ascension_level]['pickFloorSum'] += int(choice['floor'])
                    if victory:
                        card_data_map[card][ascension_level]['winCnt'] += 1

    pass


def get_raw_card_name(name):
    if '+' in name:
        return name[:name.index('+')]
    return name


def set_card_data_default(card, ascension):
    card_data_map.setdefault(card, {})
    card_data = card_data_map[card]
    if ascension not in card_data:
        card_data[ascension] = {'viewCnt': 0, 'pickCnt': 0, 'winCnt': 0, 'pickFloorSum': 0}


def set_combat_data_default(ascension):
    combat_data_map.setdefault(ascension, {'victory': 0, 'lose': 0, 'loseLayerSum': 0,
                                           'perFloor': dict((i, 0) for i in range(1, 56))})


def export_combat_data_total():
    ascensions = []
    victory = []
    lose = []
    loseLayerSum = []
    perFloor = dict((i, []) for i in range(1, 56))
    for ascension in range(0, 21):
        if ascension not in combat_data_map:
            continue
        data = combat_data_map[ascension]
        ascensions.append(ascension)
        victory.append(data['victory'])
        lose.append(data['lose'])
        for floor, l in perFloor.items():
            l.append(data['perFloor'][floor])
        if data['lose']:
            loseLayerSum.append(round(data['loseLayerSum'] / data['lose'], 1))
        else:
            loseLayerSum.append(0)
    export_data = {'进阶': ascensions, '胜利次数': victory, '失败次数': lose, '失败平均楼层': loseLayerSum}
    for floor, l in perFloor.items():
        export_data[floor] = l
    df = DataFrame(export_data)
    # print(df)
    df.to_excel('ascensions_data.xlsx', index=False)
    print('export to ascensions_data.xlsx success')


def export_card_data_total():
    card_names = []
    viewCnt = []
    pickCnt = []
    winCnt = []
    pickFloorSum = []
    for (card, card_data) in card_data_map.items():
        card_total = {'card': card, 'viewCnt': 0, 'pickCnt': 0, 'winCnt': 0, 'pickFloorSum': 0}
        for d in card_data.values():
            card_total['viewCnt'] += d['viewCnt']
            card_total['pickCnt'] += d['pickCnt']
            card_total['winCnt'] += d['winCnt']
            card_total['pickFloorSum'] += d['pickFloorSum']
        if card_total['pickCnt']:
            card_total['pickFloorSum'] = round(card_total['pickFloorSum'] / card_total['pickCnt'], 1)
        else:
            card_total['pickFloorSum'] = 0

        if card not in card_name_map:
            continue
        # card_name_map.setdefault(card, card)
        card_names.append(card_name_map[card])
        viewCnt.append(card_total['viewCnt'])
        pickCnt.append(card_total['pickCnt'])
        winCnt.append(card_total['winCnt'])
        pickFloorSum.append(card_total['pickFloorSum'])
        # print(card_total)
    export_data = {'卡牌名称': card_names, '掉落次数': viewCnt, '获取次数': pickCnt, '平均获取楼层': pickFloorSum, '获取并胜利': winCnt}
    df = DataFrame(export_data)
    # print(df)
    df.to_excel('cards_data.xlsx', index=False)
    print('export to cards_data.xlsx success')


if __name__ == '__main__':
    init_card_name_map()
    processJson()
    export_card_data_total()
    export_combat_data_total()

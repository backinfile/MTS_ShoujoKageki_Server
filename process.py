import os.path
from collections import defaultdict

import pandas
from pandas import DataFrame
import json
import openpyxl


class CardData:
    # card:{ascension:CardData}
    card_data_map = defaultdict(lambda: defaultdict(lambda: CardData()))

    def __init__(self):
        self.viewCnt = 0
        self.pickCnt = 0
        self.winCnt = 0
        self.pickFloorSum = 0
        self.singlePickCnt = 0
        self.singleWinCnt = 0
        self.firstPickFloor = 0
        self.firstSmithFloor = 0

    @staticmethod
    def process(content):
        card_cache_set = set()
        card_choices = content['event']['card_choices']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']

        for choice in card_choices:
            if int(choice['floor']) <= 0:
                continue
            for card in choice['not_picked']:
                card = get_raw_card_name(card)
                card_data = CardData.card_data_map[card][ascension_level]
                card_data.viewCnt += 1
            card = choice['picked']
            card = get_raw_card_name(card)
            if card:
                card_data = CardData.card_data_map[card][ascension_level]
                card_data.viewCnt += 1
                card_data.pickCnt += 1
                card_data.pickFloorSum += int(choice['floor'])
                if victory:
                    card_data.winCnt += 1
                if card not in card_cache_set:
                    card_cache_set.add(card)
                    card_data.singlePickCnt += 1
                    if victory:
                        card_data.singleWinCnt += 1


class CombatData:
    # ascension:{victory:cnt, lose:cnt, loseLayerSum:cnt, perFloor:{floor:cnt}}
    combat_data_map = defaultdict(lambda: CombatData())

    def __init__(self):
        self.victory = 0
        self.lose = 0
        self.loseLayerSum = 0
        self.perFloor = defaultdict(lambda: 0)

    @staticmethod
    def process(content):
        floor_reached = content['event']['floor_reached']
        ascension_level = int(content['event']['ascension_level'])
        victory = content['event']['victory']

        add_victory_data(ascension_level, floor_reached, victory)

        combat_data = CombatData.combat_data_map[ascension_level]
        if victory:
            combat_data.victory += 1
        else:
            combat_data.lose += 1
            combat_data.loseLayerSum += floor_reached
            combat_data.perFloor[floor_reached] += 1


victory_data_map = {}  # ascension:{floor:{victory:cnt, lose:cnt}}
victory_data_total = 0

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
            victory = content['event']['victory']
            ascension_level = int(content['event']['ascension_level'])
            if 'Loadout Mod' in mods:
                # print(file_name)
                continue
            CombatData.process(content)
            if floor_reached < 3:
                continue
            # if victory:
            #     if 'killed_by' in content['event']:
            #         print('killed_by' + content['event']['killed_by'])
            #     print(str(floor_reached) + " " + file_name)
            #     print(ascension_level)
            CardData.process(content)


def get_raw_card_name(name):
    if '+' in name:
        return name[:name.index('+')]
    return name


def add_victory_data(ascension, reach_floor, victory):
    global victory_data_total
    victory_data_map.setdefault(ascension, dict((i, {'victory': 0, 'lose': 0}) for i in range(1, 58)))
    if victory:
        victory_data_total += 1
        for i in range(1, reach_floor + 1):
            victory_data_map[ascension][i]['victory'] += 1
    else:
        for i in range(1, reach_floor + 1):
            victory_data_map[ascension][i]['lose'] += 1
    if ascension != -1:
        add_victory_data(-1, reach_floor, victory)


class Export:
    export_data = {}

    @staticmethod
    def export():
        Export.export_data.clear()
        Export.export_card_data_total()
        Export.export_combat_data_total()
        Export.export_victory_data()

        os.makedirs('export', exist_ok=True)
        with pandas.ExcelWriter(os.path.join('export', 'export.xlsx')) as writer:
            for sheet_name, data in Export.export_data.items():
                df = DataFrame(data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    @staticmethod
    def export_victory_data():
        ascensions = []
        perFloor = dict((i, []) for i in range(1, 58))
        for ascension in range(-1, 21):
            if ascension not in victory_data_map:
                continue
            ascensions.append(ascension)
            for floor, data in victory_data_map[ascension].items():
                victory = data['victory']
                lose = data['lose']
                if victory + lose > 0:
                    perFloor[floor].append(victory / (victory + lose))
                else:
                    perFloor[floor].append(0)
        export_data = {'进阶': ascensions}
        for floor, l in perFloor.items():
            export_data[floor] = l
        Export.export_data["各层胜率"] = export_data
        print('export to victory_data.xlsx success')

    @staticmethod
    def export_combat_data_total():
        ascensions = []
        victory = []
        lose = []
        loseLayerSum = []
        perFloor = dict((i, []) for i in range(1, 58))
        for ascension in range(0, 21):
            if ascension not in CombatData.combat_data_map:
                continue
            data = CombatData.combat_data_map[ascension]
            ascensions.append(ascension)
            victory.append(data.victory)
            lose.append(data.lose)
            for floor, l in perFloor.items():
                l.append(data.perFloor[floor])
            if data.lose > 0:
                loseLayerSum.append(round(data.loseLayerSum / data.lose, 1))
            else:
                loseLayerSum.append(0)
        export_data = {'进阶': ascensions, '胜利次数': victory, '失败次数': lose, '失败平均楼层': loseLayerSum,
                       '胜率': [v/(v+l) if v+l > 0 else 0 for v, l in zip(victory, lose)]
                       }
        for floor, l in perFloor.items():
            export_data[floor] = l
        Export.export_data["进阶胜率"] = export_data

    @staticmethod
    def export_card_data_total():
        card_names = []
        viewCnt = []
        pickCnt = []
        winCnt = []
        pickFloorSum = []
        singlePickCnt = []
        singleWinCnt = []
        for (card, card_data) in CardData.card_data_map.items():
            if card not in card_name_map:
                continue
            viewCntNum = sum((d.viewCnt for d in card_data.values()))
            pickCntNum = sum((d.pickCnt for d in card_data.values()))
            winCntNum = sum((d.winCnt for d in card_data.values()))
            singlePickCntNum = sum((d.singlePickCnt for d in card_data.values()))
            singleWinCntNum = sum((d.singleWinCnt for d in card_data.values()))
            pickFloorSumNum = sum((d.pickFloorSum for d in card_data.values()))

            card_names.append(card_name_map[card])
            viewCnt.append(viewCntNum)
            pickCnt.append(pickCntNum)
            winCnt.append(winCntNum)
            singlePickCnt.append(singlePickCntNum)
            singleWinCnt.append(singleWinCntNum)
            pickFloorSum.append(pickFloorSumNum)
            # print(card_total)
        export_data = {'卡牌名称': card_names, '掉落次数': viewCnt, '获取次数': pickCnt, '获取并胜利次数': winCnt,
                       '去重获取次数': singlePickCnt, '去重获取并胜利次数': singleWinCnt,
                       '平均获取楼层': [round(floor/pick, 1) if pick > 0 else 0 for pick, floor in zip(pickCnt, pickFloorSum)],
                       '去重胜率': [win/pick if pick > 0 else 0 for pick, win in zip(singlePickCnt, singleWinCnt)],
                       '选取率': [pick/view if view > 0 else 0 for view, pick in zip(viewCnt, pickCnt)]
                       }
        Export.export_data["卡牌数据"] = export_data


if __name__ == '__main__':
    init_card_name_map()
    processJson()
    Export.export()

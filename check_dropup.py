import csv
import logging
import argparse
import re
import os
import unicodedata

logger = logging.getLogger(__name__)


def main(args):
    drop_up_retes = [0, 5, 10]  # 今後礼装の枚数が増えてきたらここを修正すればok
    history_file = os.path.join(os.path.dirname(__file__), "history.csv")
    csv_file = open(history_file, "r", encoding="utf-8", newline="")
    f = csv.DictReader(csv_file)
    results = []
    for row in f:
        drop_up = 0
        pattern = r"【(?P<place>[\s\S]+)】(?P<num>[\s\S]+?)周(?P<num_after>.*?)\n(?P<items>[\s\S]+?)#FGO周回カウンタ"
        m = re.search(pattern, row["text"])
        if m:
            num = int(re.sub(pattern, r"\g<num>", m.group()))
            place = re.sub(pattern, r"\g<place>", m.group())
            items = re.sub(pattern, r"\g<items>", m.group())
            if place == "シャーロット ゴールドラッシュ":
                pattern_drop_up = r"塵泥UP(?P<num>[\s\S]+?)%"
                report = unicodedata.normalize("NFKC", row["text"])
                m = re.search(pattern_drop_up, report)
                if m:
                    drop_up_rate = re.sub(pattern_drop_up,
                                          r"\g<num>",
                                          m.group()).strip()
                    if drop_up_rate.isdigit():
                        if int(drop_up_rate) == 0:
                            drop_up = 0
                        else:
                            drop_up = int(drop_up_rate)
                    else:
                        drop_up = 0
                else:
                    drop_up = 0
                items = items.replace("-", "\n")
                for item in items.split("\n"):
                    if not item.startswith("塵"):
                        continue
                    else:
                        drop_num = int(item.replace("塵", ""))
                        break
                result = {"drop_up": drop_up,
                          "num_farm": num,
                          "drop_num": drop_num,
                          "id": row["id"],
                          "screen_name": row["screen_name"],
                          "date": row["time"]
                          }
                results.append(result)
    # print(results)
    # 後処理 drop_up を拾いだす
    num_farms = ["周回数"]
    dropups = {}
    for i in drop_up_retes:
        dropups[i] = [f'{i}%']
    urls = ["ソース"]
    dates = ["メモ"]
    for result in reversed(results):
        num_farms.append(result["num_farm"])
        urls.append(f'https://twitter.com/{result["screen_name"]}/status/{result["id"]}')
        dates.append(result["date"][0:10].replace('-', '/'))
        for i in drop_up_retes:
            if result["drop_up"] == i:
                dropups[i].append(result["drop_num"])
            else:
                dropups[i].append("")
    with open(args.output, 'w', encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(num_farms)
        for i in drop_up_retes:
            writer.writerow(dropups[i])
        writer.writerow(urls)
        writer.writerow(dates)


def parse_args():
    # オプションの解析
    parser = argparse.ArgumentParser(description='OpenPoseの実行')

    parser.add_argument(
                        'output',
                        )
    parser.add_argument(
                        '-l', '--loglevel',
                        choices=('warning', 'debug', 'info'),
                        default='info'
                        )
    return parser.parse_args()


if __name__ == '__main__':
    args = parse_args()
    logger.setLevel(args.loglevel.upper())
    logger.info('loglevel: %s', args.loglevel)
    lformat = '%(name)s <L%(lineno)s> [%(levelname)s] %(message)s'
    logging.basicConfig(
        level=logging.INFO,
        format=lformat,
    )
    logger.setLevel(args.loglevel.upper())

    main(args)

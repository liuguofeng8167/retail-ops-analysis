# 数据放在这个文件夹里，**不要提交到 GitHub**

## 下载

- 名称：淘宝用户行为数据集（UserBehavior）
- 地址：https://tianchi.aliyun.com/dataset/649
- 文件名：`UserBehavior.csv`（无表头）
- 体量：约 1 亿条行为记录；机器跑不动就先抽样

## 放到这里之后

```bash
# 先抽样跑通
python funnel_analysis.py --input data/UserBehavior.csv --sample 2000000 --out outputs

# 全量
python funnel_analysis.py --input data/UserBehavior.csv --out outputs
```

字段顺序：`user_id, item_id, category_id, behavior_type(pv/cart/fav/buy), timestamp`。

## 为什么不入库

1. 全量文件超过 100 MB，GitHub 单文件上限是 100 MB；
2. 数据集有使用条款，不适合再分发；
3. 按上面的地址可重新下载。

`../.gitignore` 里的 `data/*` 已经把这里的内容排除了，只保留这个说明文件。

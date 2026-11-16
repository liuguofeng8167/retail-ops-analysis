# 数据放在这个文件夹里，**不要提交到 GitHub**

## 下载

- 名称：Online Retail II
- 地址：https://archive.ics.uci.edu/dataset/502/online+retail+ii
- 文件名：`online_retail_II.xlsx`（含两个 sheet：2009-2010、2010-2011）
- 体量：约 100 万行交易明细

## 放到这里之后

```bash
python retail_analysis.py --input data/online_retail_II.xlsx --out outputs
```

或者把文件直接拖到 `运行_拖入数据文件.bat` 上。

## 为什么不入库

1. 体积大（几十 MB），超过 GitHub 的建议单文件大小；
2. 数据集有自己的使用条款，随仓库再分发不合适；
3. 任何人按上面的地址都能重新下载，入库没有额外价值。

`../.gitignore` 里的 `data/*` 已经把这里的内容排除了，只保留这个说明文件。

# 样式与工具

默认样式由`assets/default-style.json`定义，生成器按中性咨询版式创建Office文件，不携带真实客户、商标图片或商业字体文件。

```json
{
  "font_east_asia": "Noto Sans CJK SC",
  "font_latin": "Arial",
  "capacity": 500,
  "logos": []
}
```

私有配置可替换字体和最多两张Logo，图片路径相对于配置文件目录；也支持绝对路径。使用`--profile`显式指定，或保存为安装目录的`assets/style.local.json`以自动加载。私有配置与客户资源不要提交到公共仓库。

Word为A4，正文10.5pt、约15pt行距，Title约19pt，一级13pt、二级11pt；短标签加粗，浅色元数据表，页眉与页脚包含项目身份。Excel浅蓝灰表头、隔行浅灰、编号及日期清晰，按数据量设置行高；文本预计会超过单元格可读高度时拒绝静默截断。

默认字体只作声明，不保证在目标机器存在。应选择目标环境实际可用的中文字体；Linux可安装开放的Noto CJK字体，Windows和macOS按组织配置选择本机字体。不要因为文件声明了字体就声称渲染完全一致。

Python依赖见requirements.txt。`--recalculate`和`--preview`额外依赖PATH中的`soffice`或`libreoffice`。脚本不安装或下载Office，不处理音频，不调用任何外部模型API，也不上传材料。

生成器会保留工作表中的品牌图片，打印时重复标题行和编号列，并提供文字页眉。图片跨页、复杂长文分页及Office与LibreOffice差异仍需在实际目标软件检查；不能根据嵌入图片数量保证每一打印页都呈现同样Logo。

# ExcelValidator
用来检查Excel配置表数据的通用代码

最近一直在思考游戏配置表设置的合理性.废话少说了.下面才是重点.

缘起:
由于上次和同事一起修改配置表的时候,那个配置表是有复合列要求所填数据不能重复.但我们靠人眼都没能识别出来.估计这个时候像素眼也是没法子的了.

大部分的配置表数据是有规则要求的.如果用代码来解决的话.那就多了.上至C,下至go.都是可以的.
这里不说外部语言,只说一下EXCEL自带的VBA的缺陷.就是如果禁用了宏这个东东.那用起来就是很不方便的.

其实上面的也是废话.用自己熟悉的语言才是王道.^_^

所谓通用的Validator设计就有不少了.从Web的各个Field的设计,到WinForm/App的各个控件的使用.都是可以抽象出来的.所以这里的设计大多数也是有以前使用过的影子.

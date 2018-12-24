# wget_xl_images
下载xlsx表格的图片
## install
    pip3 install --user -r requirements.txt
## run
   -  python3 wget_images.py 数据表名称(不用带.xlsx)
      然后生成按日期的目录内图片
   -  python3 rename.py 数据表名称(不用带.xlsx)
      #需要先将xlsx的副本改名为zip，然后再解压，移动xl/media 到脚本路径下

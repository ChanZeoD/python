import requests
import os.path
from lxml import etree
import xlsxwriter

if __name__ == "__main__":
    if not os.path.exists('./pic'):
        os.mkdir('./pic')
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36'
    }
    index = 1

    # 生成Excel文件
    table = './文章数据.xls'
    wb = xlsxwriter.Workbook(table)
    ws = wb.add_worksheet('数据')
    # 设置单元格大小
    ws.set_column('A:A', 29)
    ws.set_column('B:B', 30)
    ws.set_column(2, 3, 10)
    # 设置单元格格式
    cell_format = wb.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
    # 设置标题行
    headData = ['封面图', '文章标题', '发布日期', '浏览量']
    head_format = wb.add_format({'bold': 1, 'align': 'center', 'font_name': u'微软雅黑', 'valign': 'vcenter'})
    for colnum in range(0, 4):
        ws.write(0, colnum, headData[colnum], head_format)

    # 获取页面源码
    url = 'https://www.aquanliang.com/blog/page/%d'
    for pageNum in range(1, 60):
        # 不同页码的url
        new_url = (url % pageNum)
        page_text = requests.get(url=new_url, headers=headers).text

        # 解析数据
        tree = etree.HTML(page_text)
        div_list = tree.xpath('//div[@class="_1ySUUwWwmubujD8B44ZDzy"]/span/div')

        for div in div_list:
            # 局部解析
            # 下载图片到本地文件夹并存进Excel文件
            img_src = div.xpath('./a//img/@src')[0]
            img_name = img_src.split('/')[-1]+'.jpg'
            img_data = requests.get(url=img_src, headers=headers).content
            img_path = './pic/'+img_name
            with open(img_path, 'wb') as fp:
                fp.write(img_data)
                print(img_src, '下载成功')
            ws.set_row(index, 100)
            ws.insert_image(index, 0, img_path, {'x_scale': 0.4, 'y_scale': 0.4})

            # 获取文章内容
            title = div.xpath('.//div[@class="_3_JaaUmGUCjKZIdiLhqtfr"]/text()')[0]
            date = div.xpath('.//div[@class="_3TzAhzBA-XQQruZs-bwWjE"]/text()')[0]
            views = div.xpath('.//div[@class="_2gvAnxa4Xc7IT14d5w8MI1"]/text()')[0]

            # 将文本存进Excel文件
            article = [title, date, views]
            print(index)
            for i in range(0, 3):
                print(article[i])
                ws.write(index, i+1, article[i], cell_format)
            index += 1
    wb.close()


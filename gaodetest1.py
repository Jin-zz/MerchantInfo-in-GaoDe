from urllib.parse import quote
from urllib import request
import json
import xlwt

#TODO 替换为上面申请的密钥
amap_web_key = '9916a6662584a50c5f8054a101a3ae4e'
poi_search_url = "http://restapi.amap.com/v3/place/text"
poi_boundary_url = "https://ditu.amap.com/detail/get/detail"
#from transCoordinateSystem import gcj02_to_wgs84

#TODO cityname为需要爬取的POI所属的城市名，nanning_areas为城市下面的所有区，classes为多个分类名集合. (中文名或者代码都可以，代码详见高德地图的POI分类编码表)
# cityname = '南宁'
# nanning_areas = ['青秀区','兴宁区','江南区','良庆区','邕宁区','西乡塘区','武鸣区']
cityname = '无锡'
# nanning_areas = ['锡山区','惠山区','滨湖区','梁溪区','新吴区']
nanning_areas = ['滨湖区']

# classes = ['商场', '银行']
classes = ['生活服务场所','旅行社','信息咨询中心','售票处','邮局','物流速递','电讯营业厅','事务所','人才市场','自来水营业厅','电力营业厅','美容美发店']

# 根据城市名称和分类关键字获取poi数据
def getpois(cityname, keywords):
    i = 1
    poilist = []
    while True:  # 使用while循环不断分页获取数据
        result = getpoi_page(cityname, keywords, i)
        print(result)
        result = json.loads(result)  # 将字符串转换为json
        if result['count'] == '0':
            break
        hand(poilist, result)
        i = i + 1
    return poilist


# 数据写入excel
def write_to_excel(poilist, cityname, classfield):
    # 一个Workbook对象，这就相当于创建了一个Excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet(classfield, cell_overwrite_ok=True)

    # 第一行(列标题)
    sheet.write(0, 0, 'x')
    sheet.write(0, 1, 'y')
    sheet.write(0, 2, 'name')
    sheet.write(0, 3, 'address')
    sheet.write(0, 4, 'type')
    sheet.write(0, 5, 'typecode')
    sheet.write(0, 6, 'src')


    for i in range(len(poilist)):
        location = poilist[i]['location']
        name = poilist[i]['name']
        lng = str(location).split(",")[0]
        lat = str(location).split(",")[1]
        add = poilist[i]['address']
        typ = poilist[i]['type']
        typc = poilist[i]['typecode']
        typc = typc[0:4]
        if poilist[i]['photos'] == []:
            # photos = [{"title":[],"url":"xxx"}]
            src = "../../images/mer_xxx.jpg"
        else:
            photos = poilist[i]['photos'][0]
            src = photos['url']
        # print(src)

        '''
        result = gcj02_to_wgs84(float(lng), float(lat))

        lng = result[0]
        lat = result[1]
        '''

        # 每一行写入
        sheet.write(i + 1, 0, float(lng))
        sheet.write(i + 1, 1, float(lat))
        sheet.write(i + 1, 2, name)
        sheet.write(i + 1, 3, add)
        sheet.write(i + 1, 4, typ)
        sheet.write(i + 1, 5, typc)
        sheet.write(i + 1, 6, src)


    # 最后，将以上操作保存到指定的Excel文件中
    book.save(r'' + cityname + "_" + classfield + '.xls')


# 将返回的poi数据装入集合返回
def hand(poilist, result):
    # result = json.loads(result)  # 将字符串转换为json
    pois = result['pois']
    for i in range(len(pois)):
        poilist.append(pois[i])


# 单页获取pois
def getpoi_page(cityname, keywords, page):
    req_url = poi_search_url + "?key=" + amap_web_key + '&extensions=all&keywords=' + quote(
        keywords) + '&city=' + quote(cityname) + '&citylimit=true' + '&offset=25' + '&page=' + str(
        page) + '&output=json'
    data = ''
    with request.urlopen(req_url) as f:
        data = f.read()
        data = data.decode('utf-8')
    return data


for clas in classes:
    classes_all_pois = []
    for area in nanning_areas:
        pois_area = getpois(area, clas)
        print('当前城区：' + str(area) + ', 分类：' + str(clas) + ", 总的有" + str(len(pois_area)) + "条数据")
        classes_all_pois.extend(pois_area)
    print("所有城区的数据汇总，总数为：" + str(len(classes_all_pois)))

    write_to_excel(classes_all_pois, cityname, clas)

    print('================分类：'  + str(clas) + "写入成功")


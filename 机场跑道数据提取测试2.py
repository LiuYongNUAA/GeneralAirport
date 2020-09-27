# -*- encoding:utf-8 -*-
import xlwt
from selenium import webdriver
import time
from selenium.common.exceptions import NoSuchElementException
html = 'https://ga.aopa.org.cn/web_airport/html/airport_publish/search_main.html'
next_xpath = "/html/body/div[1]/div[2]/div[4]/div/div/div[3]/a[@class='next']"
#创建表格
book=xlwt.Workbook(encoding="utf-8")
sheet=book.add_sheet('机场跑道信息',cell_overwrite_ok=True)
list=('机场名称','机场地址','所属管理局','机场类型','基准点坐标','机场运营人','联系电话','E-mail','机场所有',
      '机场官网','机场介绍','开放时间','最大使用机型','驻场企业数量','驻场企业名称','常驻飞机数量','常驻飞机型号',
      '机场管制空域','无线电导航及着陆设施','障碍物','塔台','空中交通管制服务频率','空中交通管制服务呼号','空中交通管制联系方式',
      '夜航支持','加油服务','充放电服务','机库','跑道型机场消防/等级','消防设备(跑道型)','直升机机场消防/等级','消防设备(直升机)',
      '应急救援服务','维修服务','餐饮服务','住宿服务','租车服务','餐厅推荐','酒店推荐','旅游景点推荐','航空活动介绍','周边活动介绍',
      '交通信息'
       ,'机场标高','标记牌','风向标','助航灯光','备用电源','其他设施','表面类型','停机位','PCN值','跑道号','飞行区指标','飞行规则',
       '跑道尺寸','坡度（纵坡）','磁偏角','表面类型','PCN值','升降带','跑道端识别号','磁方位角','真方位角','入口内移',
       '可用起飞滑跑距离（TORA）','可用起飞距离（TODA）','可用加速停止距离（ASDA）','可用着陆距离（LDA）','净空道','停止道','端安全区',
       '滑行道编号','表面类型','PCN值','宽度')
for i in range(0,76):
    sheet.write(i,0,list[i])

def homepage(home):
    driver.get(home)
    time.sleep(1)
    iframe = driver.find_elements_by_tag_name('iframe')[1]
    driver.switch_to.frame(iframe)
    return


def try_click(xpath):
    try:
        driver.find_element_by_xpath(xpath).click()
    except:
        time.sleep(2)
        driver.find_element_by_xpath(xpath).click()
    finally:
        time.sleep(0.5)
    return

 
# -*- coding: utf-8 -*-
def getdata():
    airport_data=[airport_name,]
    airport_element=('address','licence_num','airport_types','latitude','airport_run','contact_num','email','airport_holder','url','publicity',
             'opening_hours_start','max_type','on_site_enterprise_num','on_site_enterprise_name','resident_plane_num','resident_plane_type',
             'airport_control_airspace','radio_land_facilities','obstacles','is_control_tower','air_frequency','air_cry','air_contact_way',
             'is_night_flight_navigation','refuel','charge_discharge_supplier','hangar','fire_control_class','fire_fighting_apparatus','heliport_fire_control_class',
             'heliport_fire_fighting_apparatus','emergency_rescue_service','maintenance_service','catering_services','residential_services','car_rental',
             'restaurant_recommendation','hotel_recommend','travel_recommend','go_sky_introduce','peripheral_activities_introduce','traffic_information'
             )
    #爬取机场基础信息
    for i in airport_element:
        i= driver.find_element_by_id(i).text
        airport_data.append(i)   

    for j in range(0,43):

        sheet.write(j,int(airport_num),airport_data[j]) 



    runway=[]
    try:
        runway_num=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[8]/div[1]/span[2]').text
        runway.append(runway_num)
    except NoSuchElementException:
        runway_num='none'
     #爬取机场跑道信息
    if not(runway_num =='none'): 
        
        runway1=('airport_level','sign','wind_vane_status','zh_lamplight_status','backup_power_status','other_facilities','surface_type','gate_position_num',
             'jiping_pcn')
        for i in runway1:
            i= driver.find_element_by_id(i).text
            runway.append(i)          
        for r in range(1,5):    
            colum2=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[8]/div[2]/div[%s]/div[2]'%str(r)).text
            runway.append(colum2)
            colum4=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[8]/div[2]/div[%s]/div[4]'%str(r)).text
            runway.append(colum4)

        for t in range(3,14):
            data2=''
            for l in range(2,4):       
                try:
                    data=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[8]/div[%s]/div[%s]'%(str(t),str(l))).text
                
                except NoSuchElementException:
                    data='*'
                data2+=str(data)+'、'    
            runway.append(data2)
            
        try:
            for i in range(3,12):
                data=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[9]/div[%s]/div[1]'%str(i)).text
        except NoSuchElementException:
            for k in range(1,5):
                data1=''
                for j in range(3,i):
                    data=driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[3]/div[9]/div[%s]/div[%s]'%(str(j),str(k))).text           
                    data1 += str(data)+'、'
                runway.append(data1)
     #如果没有跑道信息，则用**填充所有项
    else:
        for j in range(0,33):
            data='**'
            runway.append(data)
        
    for j in range(43,76):

        sheet.write(j,int(airport_num),runway[j-43]) 
    book.save('%s.xls'%str(total_num))
    
driver = webdriver.Ie()
homepage(html)
driver.maximize_window()

# 首先定位三类机场位置：已取证通用机场（71个） 已备案机场（226个） 其他起降场地（4个）
airport_type_xpath = []
for i in range(0,3):
    airport_type_xpath.append('/html/body/div[1]/div[2]/div[2]/span[%s]' % str(i + 1))

for airport_xpath in airport_type_xpath:
    # 进入第i类机场
    try_click(airport_xpath)

    # 获取该类机场总数
    total_num_xpath = airport_xpath + '/strong'
    total_num = driver.find_element_by_xpath(total_num_xpath).text
    print("机场总数：%s" % total_num)
    
  
    # 定位第几页第几个机场
    for i in range(62,int(total_num)):
        airport_num = i + 1
        if airport_num % 20 == 0:
            page = airport_num // 20
            num = 20
        else:
            page = airport_num // 20 + 1
            num = airport_num % 20
            
        # 跳转到对应页数
        
        for j in range(page-1):
            try_click(next_xpath)

            
        ele_xpath = '/html/body/div[1]/div[2]/div[4]/div/div/table/tbody/tr[%s]/td[1]' % str(num)
        airport_name = driver.find_element_by_xpath(ele_xpath).text
        try_click(ele_xpath)
        time.sleep(1)
        getdata()
        # 这里插入数据抓取函数

        print("已读取第%d个机场：%s" % (airport_num, airport_name))
        # 返回首页
        homepage(html)
        try_click(airport_xpath)

        
    
    print('该类机场已爬取完毕')


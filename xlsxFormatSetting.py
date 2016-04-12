# -*- coding: utf-8 -*-

import sys

reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

server_ip = '120.26.38.147'
server_user = 'minkedong'
server_passwd = 'mkd'
server_port = 22

sender = 'lei.wang@mathartsys.com'
receiver = ['peter.zhang@mathartsys.com', 'huan.yang@mathartsys.com', 'lei.wang@mathartsys.com']
smtpserver = 'smtp.qiye.163.com'
username = 'lei.wang@mathartsys.com'
password = 'Walle19930215'

filePath = 'D:\\project\\project-carNew\\priceReport'
# filePath = 'D:\\work\\priceReport'
linPath = '/home/minkedong/py_env/Malibu/daily_data'

markValue = {'默认值': 2000,
             '创酷': {'2016款 1.4T 手动两驱舒适天窗版': 8000, 
                      '2016款 1.4T 自动两驱豪华型': 12000,
                      '2016款 1.4T 自动两驱舒适天窗版':10000,
                      '2016款 1.4T 自动四驱旗舰型':0},
             '乐风RV': {'2016款 1.5L 手动畅行版':2000,
                        '2016款 1.5L 自动畅行版':2000,
                        '2016款 1.5L 自动智行版':2000,
                        '2016款 1.5L 自动趣行版':2000},
             '经典科鲁兹': {'2015款 1.5L 经典 SE AT':14000,
                            '2015款 1.5L 经典 SE MT':14000,
                            '2015款 1.5L 经典 SL MT':12000},
             '迈锐宝XL': {'2016款 1.5T 双离合锐驰版':0,
                          '2016款 1.5T 双离合锐享版':0,
                          '2016款 1.5T 双离合锐尚版':0,
                          '2016款 1.5T 双离合锐耀版':0,
                          '2016款 2.5L 自动锐尚版':0,
                          '2016款 2.5L 自动锐尊版':0},
             '迈锐宝': {'2016款 1.6T 自动舒适版':27000,
                        '2016款 1.6T 自动豪华版':27000,
                        '2016款 2.0L 自动舒适版':25000,
                        '2016款 2.0L 自动豪华版':27000,
                        '2016款 2.4L 自动豪华版':27000,
                        '2016款 2.4L 自动旗舰版':27000},
             '全新科鲁兹': {'2016款 1.4T DCG豪华版':16000,
                            '2016款 1.4T DCG旗舰版':0,
                            '2016款 1.5L 手动精英版':16000,
                            '2016款 1.5L 手动时尚版':0,
                            '2016款 1.5L 自动豪华版':16000,
                            '2016款 1.5L 自动时尚天窗版':16000},
            }
markvalueColor = 'FFC7CE'

fieldList = ['source', 'price', 'promotion', 'promotion_url', 'p_date', 'agency_name', 'sales_area', 'province', 'city',
             'carseries_name', 'b_name', 'mb_name', 'cartype_name', 'total_price', 'bare_price', 'purchase_tax',
             'insurance', 'use_tax', 'card_price', 'old_price', 'compulsoryInsurance_price', 'crawl_date',
             'promotion_price', 'p_level', 'model_year', 'source2', 'MSRP', 'gap', 'source_id', 'mac', 'is_4s',
             'dealerName']

carseries = {'创酷': '<',
			 '乐风RV': '<',
			 '经典科鲁兹': '<',
			 '迈锐宝XL': '>',
			 '迈锐宝': '<',
			 '全新科鲁兹': '<'}

websiteDealer = ['汽车之家（经销商入口）', '汽车之家（报价入口）','易车（经销商入口）', '易车（报价入口）',
                 '太平洋汽车', '爱卡汽车', '新浪汽车', '搜狐汽车', '凤凰网']

area = ['雪佛兰1区', '雪佛兰2区', '雪佛兰3区', '雪佛兰4区', '雪佛兰5区', '雪佛兰6区', '雪佛兰7区', '雪佛兰8区']

province = {'雪佛兰1区': ['黑龙江', '吉林', '辽宁', '山东'],
            '雪佛兰2区': ['北京', '河北', '内蒙古', '山西', '天津'],
            '雪佛兰3区': ['湖北', '湖南', '江西'],
            '雪佛兰4区': ['福建', '广东', '广西', '海南'],
            '雪佛兰5区': ['甘肃', '河南', '宁夏', '青海', '陕西', '新疆'],
            '雪佛兰6区': ['贵州', '四川', '西藏', '云南', '重庆'],
            '雪佛兰7区': ['上海', '浙江'],
            '雪佛兰8区': ['安徽', '江苏']}

width = {'报表说明': [8.38, 25.88, 27.63, 27.63, 33.63, 12.63, 8.38, 8.38],
         '网站经销商数量变化分析': [8.38, 17.88, 11.88, 11.88, 11.88, 11.88, 11.88, 11.88, 11.88, 11.88],
         '网站报价均值分析': [8.38, 35.88, 18.25, 10.75, 10.75, 10.75, 10.75, 10.75, 10.75, 10.75, 10.75],
         '大区报价分析': [8.38, 20.88, 12.38, 12.38, 12.38, 12.38, 12.38, 12.38, 12.38, 12.38, 12.38],
         '省份报价分析': [8.38, 21, 21, 24.88, 24.88, 24.88, 24.88, 24.88, 24.88, 24.88, 24.88, 24.88, 24.88],
         '报价详细': [8.38, 8.38, 11.13, 15.13, 8.38, 11.13, 12.13, 16.13, 8.38, 16.13, 8.38, 9.13, 14.13, 13.13, 12.13, 14.13, 11.13, 9.13, 12.13, 11.13, 27.13, 12.13, 17.13, 8.38, 12.13, 9.25, 8.38, 8.38, 11.13, 8.38, 8.38, 12.13]}

style = {'标题':  {'font_family': '微软雅黑', 'font_size': 11, 'font_blod': True, 'font_color': 'FFFFFF',
                                    'patternFill_color': '002060'
                    },
         '行标题':  {'font_family': '宋体', 'font_size': 11, 'font_blod': True, 'font_color': '000000',
                                    'patternFill_color': 'D9D9D9'
                    },
         '内容':  {'font_family': '微软雅黑', 'font_size': 11, 'font_blod': False, 'font_color': '000000',
                    },
         '行标题2':  {'font_family': '宋体', 'font_size': 8, 'font_blod': True, 'font_color': '000000',
                    },
         '报表详细标题':  {'font_family': '宋体', 'font_size': 11, 'font_blod': True, 'font_color': 'FFFFFF',
                                    'patternFill_color': '5B9BD5'
                    },
         '偶数行':  {'font_family': '宋体', 'font_size': 11, 'font_blod': False, 'font_color': 'FFFFFF',
                                    'patternFill_color': 'DDEBF7'
                    },
         '奇数行':  {'font_family': '宋体', 'font_size': 11, 'font_blod': False, 'font_color': 'FFFFFF',
                    }
         }

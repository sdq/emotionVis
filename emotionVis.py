# -*- coding: utf-8 -*-
from bokeh.io import output_file, show
from bokeh.charts import Area, defaults, Donut
from bokeh.models import GeoJSONDataSource
from bokeh.plotting import figure, ColumnDataSource
from bokeh.models import Range1d, HoverTool, CustomJS
from bokeh.layouts import widgetbox,column,row, gridplot, layout
from bokeh.models.widgets import Dropdown, TextInput, Button, RadioButtonGroup, Select, Slider, DataTable, DateFormatter, TableColumn
from bokeh.embed import components
from bokeh.models.formatters import DatetimeTickFormatter, TickFormatter, String, List
import numpy as np
import pandas as pd

allData = pd.read_excel(open('reviews201612.xlsx', 'rb'), sheetname='Sheet1')

#user review data
userArray = allData["user"]
reviewArray = allData["review"]
count = len(reviewArray)
levelArray = allData["level"]
tasteArray = allData["taste"]
environmentArray = allData["environment"]
serviceArray = allData["service"]
dateArray = allData['date']

#sentiment analysis
polarityArray = allData['polarity']
magnitudeArray = allData['magnitude']
magnitudeArray = np.multiply(magnitudeArray,100)
confidenceArray = allData['confidence']

#emotion detail
surpriseArray = allData['surprise']
likeArray = allData['like']
exciteArray = allData['excite']
disappointArray = allData['disappoint']
angryArray = allData['angry']
hateArray = allData['hate']
emotionIndexArray = []
emotionNameArray = []
colorArray = ["Coral","Crimson","DarkViolet","LightSeaGreen","LightSkyBlue","LightSlateGray"]
emotionArray = ["惊喜","喜爱","兴奋","失望","愤怒","厌恶"]
for i in xrange(0,count):
    temp = [surpriseArray[i],likeArray[i],exciteArray[i],disappointArray[i],angryArray[i],hateArray[i]]
    index = temp.index(max(temp))
    emotionIndexArray.append(index)
    emotionNameArray.append(emotionArray[index])

#percentage chart
surpriseSum = sum(surpriseArray)
likeSum = sum(likeArray)
exciteSum = sum(exciteArray)
disappointSum = sum(disappointArray)
angrySum = sum(angryArray)
hateSum = sum(hateArray)
distribution = [surpriseSum/count, likeSum/count, exciteSum/count, disappointSum/count, angrySum/count, hateSum/count]
distributionLeft = [0,
                surpriseSum/count,
                surpriseSum/count + likeSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count + disappointSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count + disappointSum/count + angrySum/count]
distributionRight = [surpriseSum/count,
                surpriseSum/count + likeSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count + disappointSum/count,
                surpriseSum/count + likeSum/count + exciteSum/count + disappointSum/count + angrySum/count,
                1]


distributionSource = ColumnDataSource(
        data = dict(
            emotion = emotionArray,
            distribution = distribution,
            distributionRight = distributionRight,
            distributionLeft = distributionLeft,
            color = colorArray
        )
    )

distributionhover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">情绪 @emotion</span>
            </div>
            <div>
                <span style="font-size: 15px;">分布 @distribution</span>
            </div>
        </div>
        """
    )

distributionChart = figure(title = "情绪分布图", toolbar_location="below",
           toolbar_sticky=False, plot_width=730, plot_height=120, responsive=True,  tools = [distributionhover,"wheel_zoom","box_zoom","reset","save"])
distributionChart.hbar(y=50, height=100, left="distributionLeft",
       right="distributionRight", color="color", source = distributionSource)
distributionChart.toolbar.logo = None
distributionChart.x_range = Range1d(0, 1)

#line chart

timelineSource = ColumnDataSource(
        data = dict(
            date = dateArray,
            review = reviewArray,
            polarity = polarityArray,
            magnitude = magnitudeArray
        )
    )

timelinehover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">情感得分 @polarity</span>
            </div>
            <div>
                <span style="font-size: 15px;">情感浓度 @magnitude</span>
            </div>
        </div>
        """
    )

timelineChart = figure(title = "时间轴 - 极性与浓度", toolbar_location="below",
           toolbar_sticky=False, plot_width = 730, plot_height = 300, responsive=True, tools = [timelinehover,"wheel_zoom","box_zoom","reset","save"])
timelineChart.toolbar.logo = None
timelineChart.xaxis.formatter=DatetimeTickFormatter(
        hours=["%d %B %Y"],
        days=["%d %B %Y"],
        months=["%d %B %Y"],
        years=["%d %B %Y"])
timelineChart.line("date","polarity", color = "green", line_width=1, source = timelineSource)
timelineChart.line("date","magnitude", color = "gray", line_width=1, source = timelineSource)
timelineChart.y_range = Range1d(0, 100)

#area chart
areadata = dict(
    surprise=surpriseArray,
    like=likeArray,
    excite=exciteArray,
    disappoint=disappointArray,
    angry=angryArray,
    hate=hateArray
)

area = Area(areadata, title="情感走势图", legend="top_left",
             stack=True, xlabel='时间', ylabel='百分比', plot_width=730, plot_height=300, responsive=True, toolbar_location="below",
           toolbar_sticky=False,tools = ["wheel_zoom","box_zoom","reset","save"])
area.toolbar.logo = None
area.x_range = Range1d(0, 152)
area.y_range = Range1d(0, 1)

#scatter chart

scatterhover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">极性 @polarity</span>
            </div>
            <div>
                <span style="font-size: 15px;">浓度 @magnitude</span>
            </div>
            <div>
                <span style="font-size: 12px;">@review</span>
            </div>
        </div>
        """
    )

scatter = figure(title = "评论分析", responsive=True, plot_width=730, plot_height=730,tools = [scatterhover, "pan","wheel_zoom","box_zoom","reset","save"])
scatter.xaxis.axis_label = '情感极性'
scatter.yaxis.axis_label = '情感浓度'
scatter.y_range = Range1d(0, 100)
scatter.x_range = Range1d(0, 100)
colors = [colorArray[x] for x in emotionIndexArray]

source = ColumnDataSource(
        data = dict(
            polarity = polarityArray,
            magnitude = magnitudeArray,
            color = colors,
            legend = emotionNameArray,
            review = reviewArray
        )
    )

scatter.circle('polarity', 'magnitude',
         color='color', legend = 'legend', fill_alpha=0.2, size=10, source = source)
scatter.toolbar.logo = None

#emotion distribution chart


#rank chart
surpriseRank = np.argsort(-surpriseArray)
likeRank = np.argsort(-likeArray)
exciteRank = np.argsort(-exciteArray)
disappointRank = np.argsort(-disappointArray)
angryRank = np.argsort(-angryArray)
hateRank = np.argsort(-hateArray)

surpriseSource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in surpriseRank[0:10]],
            surprise = [surpriseArray[x] for x in surpriseRank[0:10]],
            review = [reviewArray[x] for x in surpriseRank[0:10]],
            y = range(10, 0, -1)
        )
    )

surprisehover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">惊喜占比 @surprise</span>
            </div>
        </div>
        """
    )

surpriseRankChart = figure(title = "惊喜排行榜", plot_width=365, plot_height=365, responsive=True, tools = [surprisehover,"wheel_zoom","box_zoom","reset","save"])
surpriseRankChart.hbar(y="y", height=0.5, left=0,
       right="surprise", color=colorArray[0], source = surpriseSource)
surpriseRankChart.toolbar.logo = None
#surpriseRankChart.yaxis.formatter = TickFormatter(labels={1 : "Apple", 2 : "Orange" , 3: "fuycj"})

likeSource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in likeRank[0:10]],
            like = [likeArray[x] for x in likeRank[0:10]],
            review = [reviewArray[x] for x in likeRank[0:10]],
            y = range(10, 0, -1)
        )
    )

likehover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">喜爱占比 @like</span>
            </div>
        </div>
        """
    )

likeRankChart = figure(title = "喜爱排行榜", plot_width=365, plot_height=365, responsive=True, tools = [likehover,"wheel_zoom","box_zoom","reset","save"])
likeRankChart.hbar(y="y", height=0.5, left=0,
       right="like", color=colorArray[1], source = likeSource)
likeRankChart.toolbar.logo = None

exciteSource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in exciteRank[0:10]],
            excite = [exciteArray[x] for x in exciteRank[0:10]],
            review = [reviewArray[x] for x in exciteRank[0:10]],
            y = range(10, 0, -1)
        )
    )

excitehover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">兴奋占比 @excite</span>
            </div>
        </div>
        """
    )

exciteRankChart = figure(title = "兴奋排行榜", plot_width=365, plot_height=365, responsive=True, tools = [excitehover,"wheel_zoom","box_zoom","reset","save"])
exciteRankChart.hbar(y="y", height=0.5, left=0,
       right="excite", color=colorArray[2], source = exciteSource)
exciteRankChart.toolbar.logo = None

disappointSource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in disappointRank[0:10]],
            disappoint = [disappointArray[x] for x in disappointRank[0:10]],
            review = [reviewArray[x] for x in disappointRank[0:10]],
            y = range(10, 0, -1)
        )
    )

disappointhover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">失望占比 @disappoint</span>
            </div>
        </div>
        """
    )

disappointRankChart = figure(title = "失望排行榜", plot_width=365, plot_height=365, responsive=True, tools = [disappointhover, "pan","wheel_zoom","box_zoom","reset","save"])
disappointRankChart.hbar(y="y", height=0.5, left=0,
       right="disappoint", color=colorArray[3], source = disappointSource)
disappointRankChart.toolbar.logo = None

angrySource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in angryRank[0:10]],
            angry = [angryArray[x] for x in angryRank[0:10]],
            review = [reviewArray[x] for x in angryRank[0:10]],
            y = range(10, 0, -1)
        )
    )

angryhover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">愤怒占比 @angry</span>
            </div>
        </div>
        """
    )

angryRankChart = figure(title = "愤怒排行榜", plot_width=365, plot_height=365, responsive=True, tools = [angryhover,"wheel_zoom","box_zoom","reset","save"])
angryRankChart.hbar(y="y", height=0.5, left=0,
       right="angry", color=colorArray[4], source = angrySource)
angryRankChart.toolbar.logo = None

hateRankSource = ColumnDataSource(
        data = dict(
            user = [userArray[x] for x in hateRank[0:10]],
            hate = [hateArray[x] for x in hateRank[0:10]],
            review = [reviewArray[x] for x in hateRank[0:10]],
            y = range(10, 0, -1)
        )
    )

hatehover = HoverTool(
        tooltips="""
        <div>
            <div>
                <span style="font-size: 15px;">用户 @user</span>
            </div>
            <div>
                <span style="font-size: 15px;">评论 @review</span>
            </div>
            <div>
                <span style="font-size: 12px;">厌恶占比 @hate</span>
            </div>
        </div>
        """
    )

hateRankChart = figure(title = "厌恶排行榜", plot_width=365, plot_height=365, responsive=True, tools = [hatehover,"wheel_zoom","box_zoom","reset","save"])
hateRankChart.hbar(y="y", height=0.5, left=0,
       right="hate", color=colorArray[5], source = hateRankSource)
hateRankChart.toolbar.logo = None


output_file("emotionVis.html", title="emotionVis")
l = layout([ [distributionChart],[timelineChart],
    [area],[scatter],[surpriseRankChart, likeRankChart], [exciteRankChart, disappointRankChart], [angryRankChart, hateRankChart]])
show(l)
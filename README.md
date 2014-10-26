营会分房程序2.0
=============

/*
 *  Copyright 2014.9.18  Jing Tang <tangjing725@ccnbg.org>
 *  GNU General Public License <http://www.gnu.org/licenses/>.
 *  本程序是利用google app script api 和 google drive service 实现的营会分房程序。
 *  在通知原作者的前提下，你可以使用传播以及修改本程序，但禁止用于任何商业用途。
 */


Specification:
--------------

- container：
 - 总表
 - 房间占用overview表

- input：总表上的人员信息，mark condition （有特殊房间要求的家庭，个人）, 空为随机

- Requirement:

  1. 按性别随机分配房间，从6人间开始recursiv
    - 如果年龄 45岁 以上从4人间开始recursiv
    - 某人要求要某种类型的房间 （非家庭，也不需要和谁在一起，只要求房间类型）
    - 无电梯的2个4人间（850，851）不分配给45岁以上的,同样也不分配给家庭

  2. 指定要在一个房间的人员（这里包括家庭和特殊要求在一个房间的人员）
    - 家庭
      - 无需区分性别，mark condition格式 F(2/3/4/5/6)xxx, 
        - F代表家庭间，2/3/4/5/6 代表房间类型，
        - xxx是任意字符串，不同家庭必须不同，以便区分。
        - e.g. A家庭 F2xxs: A家庭分2人间。B家庭： F4zsh： B家庭分4人间
      - 3岁以下无床 无名卡 skip
      -  非家庭但要求在一个房间的集体： 业务逻辑同家庭，所以程序并不需要区分。总表mark condition同家庭: F(2/3/4/5/6)xxx

  3. 日营
    - 分房时skip 但需要名卡 （所以只要参加营会的人，都要保留信息在总表以便打印名卡）

  4. 特殊指定要某一个房间的人员 比如牧师，讲员
    -  总表的mark condition 里写上他们预定的房间号就好，程序会自动把预定的房间分给他们

- output:
 - 总表每个人得到房间号
 - overview表上会显示房间及床铺占用情况（分到房间的人名会被填上）
 - 读取总表上的信息，批量生成名卡，pdf格式，便于打印

Resources:
----------

* 分房程序模板以及源代码：https://github.com/sijitang/ccnbg
* API: https://developers.google.com/apps-script/reference/spreadsheet/
* Limitation: https://developers.google.com/apps-script/guides/services/quotas
* 2014年南德福音营数据统计：https://sites.google.com/site/gospelcampstatistics/


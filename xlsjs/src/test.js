const xlxs = require('xlsx');
const {readFile, writeFile} = require('fs').promises;
const {utils} = xlxs;

(async function (params) {

    // 获取数据
    const excelBuffer = await readFile('./小店菜场明细.xlsx');

    // 解析数据
    const result = xlxs.read(excelBuffer, {
        type: 'buffer',
        cellHTML: false,
    });

    // console.log('TCL: result', result.Sheets['结算明细']);
    let data = result.Sheets['结算明细'];
    // console.log(data);
    // result.Sheets就是获取到的表,下层就是每个单元格了


    //搞一个分类函数,用于去重筛选的,col是列的名字,A,B,C等等
    function getUniqueArray(col) {
        let arr = [];
        for (var item in data) {
            if (item.indexOf(col) !== -1 && arr.indexOf(data[item].v) === -1) {
                arr.push(data[item].v);
            }
        }
        return arr;
    }

    //创建筛选人员和地址的数组
    let persons = getUniqueArray('Y');
    let adress = getUniqueArray('AA');

    //创建两张工作表
    let personSheet = [];
    let addressSheet = [];

    //写一个包含某个人名的统计数函数
    function getSomeMoney(index, name, data) {

        //指明哪些位置需要做统计,如果是连续的
        let sum2 = [];
        for (let i = 0; i < 22; i++) {
            sum2.push(0);
        }

        //如果是非连续的
        let sum = 0;
        for (var item in data) {
            if (data[item].v === name) {
                let str = item.slice(1);
                sum += data['X' + str].v;

                //TODO:需要处理的异常数据
                for (let i = 0; i < sum2.length && str != '1'; i++) {
                    let para = String.fromCharCode(67 + i) + str
                    if (data[para] === undefined) {
                        data[para] = {v: 0}
                    }
                    ;
                    let _item = data[para];
                    sum2[i] += _item.v;
                }
            }
        }
        //选择输出这些数据
        sum2.push(name);
        console.log(index + ":" + sum2);
        return sum2;
    }

    persons.forEach((item, index) => {
        let temp = getSomeMoney(index, item, data);
        personSheet.push(temp);
    })
    console.log("-----------------------------------------------------------------------------------------")

    //写一个包含某个小区名字的金额数函数
    function getAdsMoney(index, name, data) {
        let sum = 0;
        let sum2 = [];
        for (let i = 0; i < 22; i++) {
            sum2.push(0);
        }
        for (var item in data) {
            if (data[item].v === name) {
                let str = item.slice(2);
                sum += data['X' + str].v;
                for (let i = 0; i < sum2.length && str != '1'; i++) {
                    let para = String.fromCharCode(67 + i) + str
                    if (data[para] === undefined) {
                        data[para] = {v: 0}
                    }
                    ;
                    let _item = data[para];
                    sum2[i] += _item.v;
                }
            }
        }
        sum2.push(name);
        console.log(index + ":" + sum2);
        return sum2;
    }

    adress.forEach((item, index) => {
        let temp = getAdsMoney(index, item, data);
        addressSheet.push(temp);
    })
    console.log(personSheet, addressSheet);


    //导出文件
    const workBook2 = utils.book_new();
    const workSheet = utils.aoa_to_sheet(personSheet, {
        cellDates: true,
    });
    const workSheet2 = utils.aoa_to_sheet(addressSheet, {
        cellDates: true,
    });

// 向工作簿中追加工作表
    utils.book_append_sheet(workBook2, workSheet, '人员统计');
    utils.book_append_sheet(workBook2, workSheet2, '地址统计');

// 浏览器端和node共有的API,实际上node可以直接使用xlsx.writeFile来写入文件,但是浏览器没有该API
    const result2 = xlxs.write(workBook2, {
        bookType: 'xlsx', // 输出的文件类型
        type: 'buffer', // 输出的数据类型
        compression: true // 开启zip压缩
    });

// 写入文件
    writeFile('./输出.xlsx', result2)
        .catch((error) => {
            console.log(error);
        });

})();



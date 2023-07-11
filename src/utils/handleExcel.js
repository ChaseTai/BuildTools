// const XLSX = require('xlsx');
import * as XLSX from 'xlsx'

/**
 * Excel 操作
 * @param Object
 * tag:             标识（TradingMarket:上期所期货交易行情 UpdateGroup:更新组 DailyPaper:上期所仓单日报 Weekly:上期所库存周报）
 * metalList:       金属类别列表 上期所期货交易行情
 * currentNav:      当前nav类型 上期所期货交易行情
 * exportList:      导出功能接收数组
 * excelFiles:      读取到的excel件
 * modalList:       接收并显示excel数据的数组
 * columns:         更新组审核自定义表头
 * updateArray:     更新组表格数组
 */
// 切换tab获取对应excel内容
function getExcelData(Object){
    const fileReader = new FileReader();
    fileReader.onload = ev => {
        try {
            const data = ev.target.result;
            //将文本内容转换为二进制
            const workbook = XLSX.read(data, {
                type: "binary"
            });
            let modalList = [];
            let excelList = []; //清空接收数据
            if (Object.tag == 'TradingMarket') {
                for (let i = 0; i < Object.metalList.length; i++) {
                    //特殊处理：期货行情数据含多个金属类型，添加指定个数数组，分别装载对应类型数据
                    excelList.push([]);
                    //根据金属名称获取excel文件里同表名的数据
                    let ws = XLSX.utils.sheet_to_json(workbook.Sheets[Object.metalList[i]['metalName']]);
                    //编辑数据
                    for (let j = 0; j < ws.length; j++) {
                        excelList[i].push(ws[j]);
                    }
                    let arr = [];
                    //将解析到的数据转为表格可识别的形式
                    excelList[i].map((v, idx) => {
                        let obj = {
                            idx: idx,
                            metalType: String(v['商品名称']).trim() == '' ? '-' : v['商品名称'],
                            deliveryMonth: String(v['交割月份']).trim() == '' ? '-' : v['交割月份'],
                            previousDaySettlementPrice: String(v['前结算']).trim() == '' ? '-' : v['前结算'],
                            openingPrice: String(v['今开盘']).trim() == '' ? '-' : v['今开盘'],
                            highestPrice: String(v['最高价']).trim() == '' ? '-' : v['最高价'],
                            lowestPrice: String(v['最低价']).trim() == '' ? '-' : v['最低价'],
                            closingPrice: String(v['收盘价']).trim() == '' ? '-' : v['收盘价'],
                            settlementReferencePrice: String(v['结算参考价']).trim() == '' ? '-' : v['结算参考价'],
                            upsDownsOne: String(v['涨跌1']).trim() == '' ? '-' : v['涨跌1'],
                            upsDownsTwo: String(v['涨跌2']).trim() == '' ? '-' : v['涨跌2'],
                            dealInterest: String(v['成交手']).trim() == '' ? '-' : v['成交手'],
                            dealAmount: String(v['成交额']).trim() == '' ? '-' : v['成交额'],
                            openInterest: String(v['持仓手']).trim() == '' ? '-' : v['持仓手'],
                            changeInterest: String(v['变化']).trim() == '' ? '-' : v['变化'],
                        };
                        arr.push(obj);
                    })
                    //因金属行情含有多中金属，所以用exportList接收完整的excel内容
                    Object.exportList[i] = arr;
                }

                let index = Object.metalList.findIndex(v=>v.metalName==Object.currentNav);
                //明确当前显示内容
                modalList = Object.exportList[index];
            } else {
                let wsname = workbook.SheetNames[0]; //取第一张表，wb.SheetNames[0]是获取Sheets中第一个Sheet的名字
                let ws = XLSX.utils.sheet_to_json(workbook.Sheets[wsname]); //生成json表格内容，wb.Sheets[Sheet名]获取第一个Sheet的数据
                console.log(ws)
                //编辑数据
                // for (let i = 0; i < ws.length; i++) {
                //     excelList.push(ws[i]);
                // }
                // let arr = [];
                // //将解析到的数据转为表格可识别的形式
                // excelList.map((v, idx) => {
                //     if (Object.tag == 'UpdateGroup') {
                //         let obj = {
                //             idx: idx,
                //             commodityNo: v['商品编号'],
                //             commodityName: v['商品名称'],
                //             quoteName: v['报价名称'],
                //             quoteType: v['报价形式'],
                //             quoteUnit: v['报价单位'],
                //             marketName: v['市场'],
                //             regionName: v['地区'],
                //             warehouseName: v['仓库'],
                //             invoiceTax: v['发票'],
                //             quoteRemark: v['备注'],
                //         };
                //         arr.push(obj);
                //     } else if (Object.tag == 'DailyPaper') {
                //         let obj = {
                //             idx: idx,
                //             metalType: v['金属'],
                //             registerStockQuantity: v['期货'],
                //             changeStockQuantity: v['增减'],
                //         };
                //         arr.push(obj);
                //     } else if (Object.tag == 'Weekly') {
                //         let obj = {
                //             idx: idx,
                //             metalType: v['金属'],
                //             totalStockQuantity: v['本周库存(小计)'],
                //             totalFuturesQuantity: v['本周库存(期货)'],
                //             changeStockQuantity: v['库存增减(小计)'],
                //             changeFuturesQuantity: v['库存增减(期货)'],
                //             totalCapacityQuantity: v['可用库容(本周)'],
                //             changeCapacityQuantity: v['可用库容(增减)'],
                //         };
                //         arr.push(obj);
                //     } else if (Object.tag == 'SensitiveWord') {
                //         let obj = {
                //             idx: idx,
                //             sensitiveWord: v['过滤词名称'],
                //         };
                //         arr.push(obj);
                //     } else if (Object.tag == 'UpdateGroupUpdate') {
                //         let columnsKey = [];
                //         let obj = { idx: idx, };
                //         for (let i in v) { columnsKey.push(i); }
                //         columnsKey.forEach(sub=>{
                //             let c = Object.columns.find(i=>i.label==sub);
                //             if (c) {
                //                 obj[c.prop] = v[c.label]
                //             }
                //         })
                //         arr.push(obj);
                //     }
                // })
                //表格应显示内容
                // modalList = arr;
            }
            //使用回调接收“表格应显示内容”，使表格正常显示；注意：直接传值后赋值，无法正常显示表格内容
            if (Object.callback) {
                Object.callback(modalList);
            }
        } catch (e) {
            console.log(e);
        }
    };
    fileReader.readAsBinaryString(Object.excelFiles);
}

// 下载导入模板
function downloadTemplate(fileName){
    window.location.href = `/${fileName}.xlsx`;
}

// 导出表格
/**
 * 导出表格
 * @param Object
 * tag:         标识 TradingMarket:交易行情表格 UpdateGroup: 更新组表格
 * modalList:   当前表格对应的数组对象
 * metalList:   金属类型数组
 * exportList:  解析后的数据数组
 * exportName:  导出文件的名称
 */
function exportExcel(Object){
    // if (Object.modalList.length == 0){
    //     return;
    // }
    //新建工作簿
    let wb = XLSX.utils.book_new();
    if (Object.tag == 'TradingMarket') {
        //交易行情含多种金属数据，因此每种金属应独立生成表格后合并到同一excel文件中
        Object.metalList.forEach((item, index)=>{
            //生成每种金属的行情数据
            let dealInterestTotal = 0, openInterestTotal = 0, changeInterestTotal = 0;
            let data = [];
            Object.exportList[index].forEach(v=>{
                dealInterestTotal += v.dealInterest;
                openInterestTotal += v.openInterest;
                changeInterestTotal += v.changeInterest;
                let obj = {
                    商品名称: v.metalType, 交割月份: v.deliveryMonth, 前结算: v.previousDaySettlementPrice, 今开盘: v.openingPrice, 最高价: v.highestPrice, 最低价: v.lowestPrice, 收盘价: v.closingPrice
                    , 结算参考价: v.settlementReferencePrice, 涨跌1: v.upsDownsOne, 涨跌2: v.upsDownsTwo, 成交手: v.dealInterest, 成交额: v.dealAmount, 持仓手: v.openInterest, 变化: v.changeInterest
                };
                data.push(obj);
            })
            // 是否包含统计(默认不包含)
            if (Object.includeSum) {
                let sum = {商品名称: '小计', 交割月份: null, 前结算: null, 今开盘: null, 最高价: null, 最低价: null, 收盘价: null
                    , 结算参考价: null, 涨跌1: null, 涨跌2: null, 成交手: dealInterestTotal, 持仓手: openInterestTotal, 变化: changeInterestTotal};
                data.push(sum);
            }
            // let data = Object.exportList[index].map(v => {
            //     return {商品名称: v.metalType, 交割月份: v.deliveryMonth, 前结算: v.previousDaySettlementPrice, 今开盘: v.openingPrice, 最高价: v.highestPrice, 最低价: v.lowestPrice, 收盘价: v.closingPrice
            //         , 结算参考价: v.settlementReferencePrice, 涨跌1: v.upsDownsOne, 涨跌2: v.upsDownsTwo, 成交手: v.dealInterest, 持仓手: v.openInterest, 变化: v.changeInterest};
            // });
            //将金属行情JSON数据转为工作表
            let sheet = XLSX.utils.json_to_sheet(data);
            //将工作表添加到工作簿，并为工作表设置表名，表名非必要元素，默认为sheet
            XLSX.utils.book_append_sheet(wb, sheet, item.metalName);
        })
    } else if (Object.tag == 'UpdateGroup') {
        let data = Object.modalList.map(v => {
            return {报价编号: v.quotationNumber, 商品编号: v.tradeNumber, 商品名称: v.tradeName, 报价名称: v.quotationName, 报价单位: v.quotationUnit, 市场: v.market, 地区: v.regionName
                , 仓库: v.warehouse, 发票: v.invoiceTax, 状态: v.state, 备注: v.remark};
        });
        let sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, sheet);
    } else if (Object.tag == 'Log') {
        let data = Object.modalList.map(v => {
            return {访问时间: v.searchTime, 用户ID: v.searchUserId, 用户IP: v.searchUserIp, 查询词: v.searchWord};
        });
        let sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, sheet);
    } else if (Object.tag == 'UserData') {
        let data = Object.modalList.map(v => {
            return {日期: v.statisticalDate, 新增用户: v.newUserNum, 启动次数: v.startAppNum, 累计用户: v.cumulativeUserNum,
                次日留存率: v.nextDayRetentionRate, 平均单次使用时长: v.averageSingleUsageTime,
            };
        });
        let sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, sheet);
    } else if (Object.tag == 'OperationalData') {
        let data = Object.modalList.map(v => {
            return {日期: v.statisticalDate, UV数: v.clientIdNum, 独立IP数: v.clientIpNum, PV数: v.pageView, 平均停留时间: v.averageResidenceTime,
                跳出率: v.bounceRate,
            };
        });
        let sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, sheet);
    } else if (Object.tag == 'DataDetail') {
        let data = Object.modalList.map(v => {
            return {页面: v.pageName, UV数: v.clientIdNum, 独立IP数: v.clientIpNum, 平均停留时间: v.averageResidenceTime,
                跳出率: v.bounceRate,
            };
        });
        let sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, sheet);
    }
    const workbookBlod = workbook2blob(wb);
    openDownloadDialog(workbookBlod, `${Object.exportName}.xlsx`);
}

// 将workbook装化成blob对象
function workbook2blob(workbook) {
    // 生成excel的配置项
    let wopts = {
        // 要生成的文件类型
        bookType: 'xlsx',
        // // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        bookSST: true,
        type: 'binary'
    }
    let wbout = XLSX.write(workbook, wopts)
    // 将字符串转ArrayBuffer
    function s2ab(s) {
        let buf = new ArrayBuffer(s.length)
        let view = new Uint8Array(buf)
        for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff
        return buf
    }
    let blob = new Blob([s2ab(wbout)], {
        type: 'application/octet-stream'
    })
    return blob
}

// 将blob对象创建bloburl，然后用a标签实现弹出下载框
function openDownloadDialog(blob, fileName) {
    if (typeof blob == 'object' && blob instanceof Blob) {
        blob = URL.createObjectURL(blob) // 创建blob地址
    }
    let aLink = document.createElement('a')
    aLink.href = blob
    // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，有时候 file:///模式下不会生效
    aLink.download = fileName || ''
    let event
    if (window.MouseEvent) event = new MouseEvent('click')
    //   移动端
    else {
        event = document.createEvent('MouseEvents')
        event.initMouseEvent(
            'click',
            true,
            false,
            window,
            0,
            0,
            0,
            0,
            0,
            false,
            false,
            false,
            false,
            0,
            null
        )
    }
    aLink.dispatchEvent(event)
}

export {
    getExcelData,
    downloadTemplate,
    exportExcel,
    workbook2blob,
    openDownloadDialog
}

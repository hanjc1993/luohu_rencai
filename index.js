const xlsx = require("node-xlsx"),
  fs = require("fs"),
  list = xlsx.parse("./人才引进公示.xlsx");
/**
 * 功能：对数组的指定位置赋值，并填充之前的空位
 * 如果直接对地址赋值，可能产生empty值，影响遍历
 * @param { Array } arr 输入数组
 * @param { Number } lastIdx 输出数组的末位序号
 * @param { Number } lastVal 输出数组的末位值
 */
function add0(arr, lastIdx, lastVal) {
  for (let t = lastIdx - arr.length; t > 0; t--) {
    arr.push(); // 粗暴填充undefined
  }
  arr[lastIdx] = lastVal;
  return arr; // 兼容
}

function start(noMd) {
  const data = list,
    companys = {}, // 用于创建表1，公司引入情况
    sheet1Head = [], // 表1的表头
    sheetCnt = []; // 用于创建表2，月份概况
  let sheetCntTotal = 0; // 表2的总累计数量

  (function transform() {
    // 本方法主要将公示表转义为统计表的格式
    data.forEach((item, idx) => {
      let newCompany = 0; // 本批次新增的公司
      idx = Math.floor(idx / 2); // 两个"半月" 合并成一个 "整月"
      item.data.forEach((rowData) => {
        const cName = rowData[0].replace(/\s/g, "");
        const cData = companys[cName];
        if (cData) {
          if (cData[idx]) {
            cData[idx] += 1;
            return;
          }
          add0(cData, idx, 1); // 举例：对[1,2] 的第4位赋值，需要补全第3位
          return;
        }
        companys[cName] = add0([], idx, 1); // 举例：对[] 的第4位赋值，需要补全前3位
        newCompany++;
      });
      // 其他数据处理
      const sheetName = item.name;
      sheetCnt.push([sheetName, item.data.length, newCompany]);
      sheetCntTotal += item.data.length;
      sheet1Head[idx] = `${sheetName.slice(0, 2)}/${sheetName.slice(2, 4)}`;
    });
  })();

  // 下面根据累计数量排序，名称无序
  const oCompanys = [],
    monthCnt = Math.ceil(list.length / 2); // 原始表中一共有几个月的数据
  (function sortBySum() {
    let focusCnt = 0; // 用于打印log
    for (let key in companys) {
      const cMonthCnt = companys[key];
      if (monthCnt > cMonthCnt.length) {
        add0(cMonthCnt, monthCnt - 1); // 举例：可能没有后两个月的数据
      }
      let cSum = cMonthCnt.reduce((sum = 0, item = 0) => sum + item); // undefined替换为0，首次执行时sum也可能是0
      oCompanys.push([key, ...cMonthCnt, cSum]);
      // 其他数据处理
      if (cSum >= monthCnt * 2) {
        // 此处半月时不精准，稍微放宽了要求，无所谓
        focusCnt++;
      }
    }
    const lastIdx = oCompanys[0].length - 1;
    oCompanys.sort((b, a) => {
      return a[lastIdx] - b[lastIdx];
    });
    // 其他数据处理
    sheetCnt.push([
      "累计引进数量",
      sheetCntTotal,
      "家属可以随迁，实际引进人数大概为此数字的1.5-2倍",
    ]);
    sheetCnt.push(["平均每月2条及以上的公司数", focusCnt]);
  })();

  const outputName = "人才引进统计";
  // 输出
  (function output() {
    sheet1Head.unshift("公司名称");
    sheet1Head.push("累计");
    oCompanys.unshift(sheet1Head);
    sheetCnt.unshift(["", "条数", "新增公司数"]);
    const buffer = xlsx.build([
      { name: "公司名单", data: oCompanys },
      { name: "月份概况", data: sheetCnt },
    ]);

    console.log(
      `\n5月至今${outputName}共${list.length / 2}个月，按公示累计条数对公司排序`
    ); // 一个申请可能包含配偶子女等多人，目测实际人数为条数的1.65倍
    fs.writeFileSync(`./${outputName}.xlsx`, buffer);
    console.log(`统计数据已输出为“${outputName}.xlsx”\n`);
  })();

  if (noMd) {
    // 由于excel没法进行二进制比对，不知道每次改了什么
    // 所以默认输出.md格式的表格
    // .md文件不包含表二“月份概况”的内容
    return;
  }
  (function outputMd() {
    let content = "|排名|公司名称|累计|\n|-|-|-|\n"; // 添加表头
    oCompanys.forEach((rowData, idx) => {
      if (idx === 0) {
        return;
      }
      // 添加内容
      const companyName = rowData[0],
        total = rowData[rowData.length - 1] || 0;
      content += `${idx}|${companyName}|${total}|\n`; // excel中没有就空着，这里需要输出为0
    });
    fs.writeFileSync(`./${outputName}.md`, content);
    console.log(`.md 格式的统计表格已同步输出\n`);
  })();
}

const arg = process.argv.splice(2)[0]; // 只校验第一个参数
start(arg === "noMd");

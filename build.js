const ejsexcel = require("./ejsExcel")
const fs = require("fs")
const util = require("util")
const readFileAsync = util.promisify(fs.readFile)
const writeFileAsync = util.promisify(fs.writeFile)

async function test() {
    //获得Excel模板的buffer对象
    const exlBuf = await readFileAsync(`${__dirname}/public/excelFile/2.xlsx`)
    //数据源
    const data = [
        [{ location: "腾源名苑", dong: "21", danyuan: "4", hao: "2", dh: "4-604" }],
        [
            {
                content: "2021年1-6月份/112*6",
                wan: "",
                qian: "",
                bai: "6",
                shi: "7",
                ge: "2",
            },
        ],
        [
            {
                content: "2021年1-6月份/112*6",
                wan: "",
                qian: "",
                bai: "6",
                shi: "7",
                ge: "2",
            },
        ],
        [
            {
                content: "2021年1-6月份/112*6",
                wan: "",
                qian: "",
                bai: "6",
                shi: "7",
                ge: "2",
            },
        ],
        [
            {
                content: "2021年1-6月份/112*6",
                wan: "",
                qian: "",
                bai: "6",
                shi: "7",
                ge: "2",
            },
        ],
        [
            {
                content: "2021年1-6月份/112*6",
                wan: "",
                qian: "",
                bai: "6",
                shi: "7",
                ge: "2",
            },
        ],
        [
            {
                content: "（2次加压设备维修分摊费62元）",
                wan: "",
                qian: "",
                bai: "",
                shi: "6",
                ge: "2",
            },
        ],
        [
            {
                allmoney: "941",
            },
        ],
    ]
    //用数据源(对象)data渲染Excel模板
    //cachePath 为编译缓存路径, 对于模板文件比较大的情况, 可显著提高运行效率, 绝对路径, 若不设置, 则无缓存
    const exlBuf2 = await ejsexcel.renderExcel(exlBuf, data, { cachePath: __dirname + "/cache/" })
    await writeFileAsync(`${__dirname}/public/excelFile/test2.xlsx`, exlBuf2)
    console.log("生成test2.xls")
}

module.exports = test

var express = require("express")
var router = express.Router()
const ejsexcel = require("../ejsExcel")
const fs = require("fs")
const util = require("util")
const readFileAsync = util.promisify(fs.readFile)
const writeFileAsync = util.promisify(fs.writeFile)
const path = require("path")

const bigNum = {
    0: "零",
    1: "壹",
    2: "贰",
    3: "叁",
    4: "肆",
    5: "伍",
    6: "陆",
    7: "柒",
    8: "捌",
    9: "玖",
}

router.get("/", function (req, res, next) {
    res.render("index")
})

router.post("/handle", async (req, res, next) => {
    let { location, nian, yue, ri, fanghao, username, phone, unit, filename, dazhe } = req.body
    const resultList = []
    try {
        const obj = req.body
        Object.keys(obj).forEach((item) => {
            if (item.indexOf("_label") !== -1) {
                const key = item.split("_label")[1]
                !resultList[key] && (resultList[key] = {})
                Object.assign(resultList[key], {
                    text: obj[item],
                })
            }
            if (item.indexOf("_value") !== -1) {
                const key = item.split("_value")[1]
                !resultList[key] && (resultList[key] = {})
                Object.assign(resultList[key], {
                    val: eval(obj[item]) || 0,
                    valText: obj[item],
                })
            }
        })
    } catch (error) {
        res.send("提交失败,可能填写格式错误导致程序分析不出,生成不了excel" + "<br/>" + error)
    }
    let allmoney = 0
    resultList.forEach((item) => {
        allmoney += item.val
    })
    if (!dazhe) {
        dazhe = 1
    } else {
        dazhe = dazhe / 10
    }
    allmoney = allmoney * dazhe
    let jiaos = jiao(allmoney)
    if (jiaos) {
        allmoney = Math.floor(allmoney) + jiaos * 0.1
    }
    //获得Excel模板的buffer对象
    const exlBuf = await readFileAsync(
        "./importExcel/template_render.xlsx"
        // path.resolve(__dirname, "./../public/importExcel/template_render.xlsx")
    )
    //数据源
    const data = []
    try {
        data.push([
            {
                location: location,
                nian: nian,
                yue: yue,
                ri: ri,
                fanghao: fanghao,
                username: username,
                phone: phone,
                unit: unit,
            },
        ])
        resultList.forEach((item) => {
            data.push([
                {
                    content: handleContent(item.text, item.val, item.valText),
                    wan: handleCout(item.val).wan,
                    qian: handleCout(item.val).qian,
                    bai: handleCout(item.val).bai,
                    shi: handleCout(item.val).shi,
                    ge: handleCout(item.val).ge,
                    jiao: jiao(item.val),
                },
            ])
        })
        data.push([
            {
                allmoney: allmoney,
            },
        ])
        data.push([
            {
                wan: bigNum[handleCout(allmoney).wan || 0],
                qian: bigNum[handleCout(allmoney).qian || 0],
                bai: bigNum[handleCout(allmoney).bai || 0],
                shi: bigNum[handleCout(allmoney).shi || 0],
                ge: bigNum[handleCout(allmoney).ge || 0],
                jiao: bigNum[jiao(allmoney) || 0],
            },
        ])
    } catch (error) {
        res.send("处理数据失败", error)
    }
    //用数据源(对象)data渲染Excel模板
    //cachePath 为编译缓存路径, 对于模板文件比较大的情况, 可显著提高运行效率, 绝对路径, 若不设置, 则无缓存
    try {
        const exlBuf2 = await ejsexcel.renderExcel(exlBuf, data, {
            cachePath: "./cache/",
            // cachePath: path.resolve(__dirname, "./../cache/"),
        })
        await writeFileAsync(
            `./excelFile/${filename}.xlsx`,
            // path.resolve(__dirname, `./../public/excelFile/${filename}.xlsx`),
            exlBuf2
        )
    } catch (error) {
        res.send("提交失败,可能你已经打开相同命名的文件了" + "<br/>" + error)
    }
    console.log(
        `生成文件路径: ${`./excelFile/${filename}.xlsx`}`
        // `生成文件路径: ${path.resolve(__dirname, `./../public/excelFile/${filename}.xlsx`)}`
    )
    res.send("提交成功")
})

function handleContent(content, value, str) {
    if (!content) {
        return ""
    } else if (content && !value) {
        return content
    } else {
        return content + "/" + str
    }
}

function jiao(value) {
    if (value > 0) {
        return parseInt((value * 10) % 10)
    } else {
        return ""
    }
}

function handleCout(value) {
    let valueStr = Math.floor(value).toString()
    let type = ["ge", "shi", "bai", "qian", "wan"]
    let obj = {
        ge: "",
        shi: "",
        bai: "",
        qian: "",
        wan: "",
    }
    if (value > 0) {
        let index = 0
        for (let i = valueStr.length - 1; i >= 0; i--) {
            obj[type[i]] = valueStr[index] || ""
            ++index
        }
    }
    return obj
}

module.exports = router

package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx"
)

/*
* 作者:陈晓宇
* 358860528@qq.com
* 批量导出excel2json
* 支持横向导出
* 竖向导出
 */

var (
	outpath = flag.String("outpath", "./out", "outpath json path default = ./out")
)

func walkFunc(path string, info os.FileInfo, err error) error {
	//fmt.Printf("%s\n", path)
	fext := filepath.Ext(path)
	//排除打开文件
	if strings.Index(path, "~$") != -1 {
		return nil
	}
	if fext == ".xlsx" || fext == ".xls" {
		DoFile(path)
	}
	return nil
}

func main() {

	flag.Parse()
	fmt.Println(*outpath)
	filepath.Walk("./", walkFunc)
	//DoFile("VIP表.xlsx")
}

func DoFile(fname string) {
	excelFileName := fname
	xlFile, error := xlsx.OpenFile(excelFileName)
	if error != nil {
		fmt.Println(excelFileName+":", error)
	}
	fmt.Printf("开始处理excel: %s\n", excelFileName)
	for _, sheet := range xlFile.Sheets {
		fmt.Printf("文件: %s\n", sheet.Name)
		Label1, _ := sheet.Cell(0, 0).String()
		Label2, _ := sheet.Cell(1, 0).String()
		Label3, _ := sheet.Cell(2, 0).String()
		//复合标准头才处理
		if Label1 == "文件名：" && Label2 == "类  名：" && Label3 == "导出类型:" {
			fmt.Printf("开始处理: %s\n", sheet.Name)
			outname, _ := sheet.Cell(0, 1).String()
			//outclass, _ := sheet.Cell(1, 1).String()
			outtype, _ := sheet.Cell(2, 1).String()
			switch outtype {
			case "1":
				//data := []interface{}
				data := Data2Array(sheet)
				bytes, err := json.Marshal(data)
				if err == nil {
					os.MkdirAll(*outpath+"/", 0666)
					ioutil.WriteFile(*outpath+"/"+outname+".json", bytes, 0666)
				}
			case "2":
				//data := map[string]interface{}
				data := Data2Map(sheet)
				bytes, err := json.Marshal(data)
				if err == nil {
					ioutil.WriteFile(*outpath+"/"+outname+".json", bytes, 0666)
				}
			default:

			}
		}
	}
}

//数据装换成数组 横向
func Data2Array(sheet *xlsx.Sheet) (out []interface{}) {

	out = []interface{}{}
	//读取注释
	//读取类型
	//读取名字
	//sNoteS := make([]string, 0, sheet.MaxCol)
	sTypeS := make([]string, 0, sheet.MaxCol)
	sNameS := make([]string, 0, sheet.MaxCol)
	sClientS := make([]string, 0, sheet.MaxCol)
	for i := 0; i < sheet.MaxCol; i++ {

		sT, _ := sheet.Cell(4, i).String()      //类型
		sName, _ := sheet.Cell(5, i).String()   //名字
		sClient, _ := sheet.Cell(6, i).String() //值
		//sVaule := sheet.Cell(7, i)            //值
		//sNoteS = append(sNoteS)
		sTypeS = append(sTypeS, sT)
		sNameS = append(sNameS, sName)
		sClientS = append(sClientS, sClient)
	}

	for j := 7; j < sheet.MaxRow; j++ {
		//读取所有数组值
		sValue := map[string]interface{}{}
		for i := 0; i < sheet.MaxCol; i++ {

			jval := sheet.Cell(j, i)
			if i == 8 { //第一个字段为空跳过这一行
				sStr, _ := jval.String()
				if sStr == "" {
					break
				}
			}
			switch sTypeS[i] {
			case "string":
				sStr, _ := jval.String()
				sValue[sNameS[i]] = sStr
				//fmt.Println(sNameS[i], sStr)
			case "double":
				sDouble, _ := jval.Float()
				sValue[sNameS[i]] = sDouble
				//fmt.Println(sNameS[i], sDouble)
			default:
				sInt, _ := jval.Int64()
				sValue[sNameS[i]] = sInt
				//fmt.Println(sNameS[i], sInt)
			}
		}
		if len(sValue) > 0 {
			//fmt.Println(sValue)
			out = append(out, sValue)
		}

	}
	//fmt.Println(sTypeS)
	//fmt.Println(sNameS)
	//fmt.Println(out)
	//for 读取所有数据
	//处理空行
	return
}

//数据装换成数组 竖向 全局表 类型
func Data2Map(sheet *xlsx.Sheet) (out map[string]interface{}) {
	//读取注释
	//读取类型
	//读取名字
	//读取值
	out = map[string]interface{}{}
	for i := 3; i < sheet.MaxRow; i++ {
		sT, _ := sheet.Cell(i, 1).String()    //类型
		sName, _ := sheet.Cell(i, 2).String() //名字
		//sClient, _ := sheet.Cell(i, 3).String() //名字
		sVaule := sheet.Cell(i, 4) //值
		switch sT {
		case "string":
			sStr, _ := sVaule.String()
			out[sName] = sStr
		case "double":
			sDouble, _ := sVaule.Float()
			out[sName] = sDouble
		default:
			sInt, _ := sVaule.Int64()
			out[sName] = sInt
		}

	}
	return
}

package main

import (
	"fmt"
	"html/template"
	"io/ioutil"
	"net/http"
	"os"
)

// InputPath 输入文件路径
var InputPath string

// OutputPath 输出文件路径
var OutputPath string

// IP 树莓派ip
var IP string

// Port 端口
var Port string

// FileInfo 文件信息
type FileInfo struct {
	FileName    string
	HostAddress string
}

// DataResponse 数据相应
type DataResponse struct {
	Data  [][]string
	Title string
}

var boshiData [][]string
var shuoshiData [][]string
var benke17Data [][]string
var benke18Data [][]string
var benke19Data [][]string
var benke20Data [][]string

func showResult(w http.ResponseWriter, r *http.Request) {
	if r.Method == "GET" {
		t, err := template.ParseFiles("template/fileList.html")
		errCheck(err)
		var indexResponse = []FileInfo{}
		for _, fileName := range getFileList(OutputPath) {
			var fileInfo = FileInfo{
				FileName:    fileName,
				HostAddress: "http://" + IP + ":" + Port,
			}
			indexResponse = append(indexResponse, fileInfo)
		}
		errCheck(t.Execute(w, indexResponse))
	}
}

func uploadFile(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, _ := template.ParseFiles("template/upload.html")
		t.Execute(w, nil)
	} else if r.Method == "POST" {
		os.RemoveAll(InputPath)
		os.RemoveAll(OutputPath)
		os.Mkdir(OutputPath, 0777)
		os.Mkdir(InputPath, 0777)
		r.ParseForm()
		r.ParseMultipartForm(1024 * 1024 * 5)

		raw, _, err := r.FormFile("file")

		if err != nil {
			return
		}

		fileBytes, err := ioutil.ReadAll(raw)

		if err != nil {
			return
		}

		fileName := InputPath + "1" + ".xlsx"

		xlsxFile, err := os.Create(fileName)

		if err != nil {
			return
		}

		defer xlsxFile.Close()
		xlsxFile.Write(fileBytes)
		readExcel(fileName)
		filter()
		w.Header().Set("Location", "http://"+IP+":"+Port+"/fileList")
		w.WriteHeader(302)
	}
}

func queryInfo(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/index.html")
		errCheck(err)
		var fileInfo = FileInfo{
			HostAddress: "http://" + IP + ":" + Port,
		}
		t.Execute(w, fileInfo)
	}
}

func showBoshiList(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  boshiData,
			Title: "博士生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func showShuoshiList(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  shuoshiData,
			Title: "研究生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func showBenke17List(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  benke17Data,
			Title: "2017级本科生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func showBenke18List(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  benke18Data,
			Title: "2018级本科生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func showBenke19List(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  benke19Data,
			Title: "2019级本科生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func showBenke20List(w http.ResponseWriter, r *http.Request) {

	if r.Method == "GET" {
		t, err := template.ParseFiles("template/data.html")
		errCheck(err)
		var dataResponse = DataResponse{
			Data:  benke20Data,
			Title: "2020级本科生未打卡名单",
		}
		t.Execute(w, dataResponse)
	}
}

func initial() {
	InputPath = "/Users/duyi/Desktop/input/"
	OutputPath = "/Users/duyi/Desktop/output/"
	IP = "localhost"
	Port = "8888"
}

func main() {

	initial()
	files := http.FileServer(http.Dir(OutputPath))
	http.Handle("/static/", http.StripPrefix("/static/", files))
	http.HandleFunc("/uploadforsyxonly", uploadFile)
	http.HandleFunc("/", queryInfo)
	http.HandleFunc("/bs", showBoshiList)
	http.HandleFunc("/yjs", showShuoshiList)
	http.HandleFunc("/bk17", showBenke17List)
	http.HandleFunc("/bk18", showBenke18List)
	http.HandleFunc("/bk19", showBenke19List)
	http.HandleFunc("/bk20", showBenke20List)
	http.HandleFunc("/fileList", showResult)
	fmt.Println("Server started ...")
	http.ListenAndServe(IP+":"+Port, nil)
}

package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"mime/multipart"
	"net/http"
	"os"
	"strings"

	"github.com/FerdinaKusumah/excel2json"
	excelize "github.com/xuri/excelize/v2"
)

const maxUploadSize = 10 * 1024 * 1024 // 10 mb
const dir = "data/download/"

var fileName string

func convertExcelIntoJson(w http.ResponseWriter, r *http.Request) {
	defer r.Body.Close()

	if r.Method != "POST" {
		respondWithError(w, http.StatusBadRequest, "Invalid Method")
		return
	}

	// 32 MB is the default used by FormFile()
	err := r.ParseMultipartForm(32 << 20)
	checkError(w, err)

	files := r.MultipartForm.File["file"]

	if len(files) != 1 {
		respondWithError(w, http.StatusBadRequest, "Please provide only one excel file")
		return
	}

	for _, fileHeader := range files {
		fileName = fileHeader.Filename

		segments := strings.Split(fileName, ".")
		extension := segments[len(segments)-1]
		if extension != "xlsx" {
			respondWithError(w, http.StatusBadRequest, "Please provide excel file having xlsx extension")
			return
		}
		if fileHeader.Size > maxUploadSize {
			respondWithError(w, http.StatusBadRequest, fmt.Sprintf("The uploaded image is too big: %s. Please use an image less than 1MB in size", fileHeader.Filename))
			return
		}

		// Open the file
		file, err := fileHeader.Open()
		checkError(w, err)

		defer file.Close()

		buff := make([]byte, 512)
		_, err = file.Read(buff)
		checkError(w, err)

		_, err = file.Seek(0, io.SeekStart)
		checkError(w, err)

		err = os.MkdirAll(dir, os.ModePerm)
		checkError(w, err)

		f, err := os.Create(dir + fileHeader.Filename)
		checkError(w, err)

		defer f.Close()

		_, err = io.Copy(f, file)
		checkError(w, err)
	}
	//	conversion(w, fileName)
	conversion2(w, fileName)
}

func checkError(w http.ResponseWriter, err error) {
	if err != nil {
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
		return
	}
}

func respondWithJson(w http.ResponseWriter, code int, payload interface{}) {
	response, _ := json.Marshal(payload)
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(code)
	w.Write(response)
}

func respondWithError(w http.ResponseWriter, code int, msg string) {
	respondWithJson(w, code, map[string]string{"error": msg})
}

func conversion(w http.ResponseWriter, file string) {
	name, _ := fetchSheetName(file)

	var (
		result []*map[string]interface{}
		err    error
		// select sheet name
		sheetName = name
		// if you want to show all headers just passing nil or empty list
		headers = []string{}
	)

	if result, err = excel2json.GetExcelFilePath(dir+file, sheetName, headers); err != nil {
		log.Println(`unable to parse file, error: %s`, err)
		respondWithError(w, http.StatusBadRequest, fmt.Sprintf("%v", err))
	}
	respondWithJson(w, http.StatusAccepted, result)
}

func main() {
	http.HandleFunc("/convert-excel-to-json", convertExcelIntoJson)
	http.HandleFunc("/convert-csv-to-json", csvToJson)
	log.Println("Service Started at 8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}

func csvToJson(w http.ResponseWriter, r *http.Request) {

	// 32 MB is the default used by FormFile()
	err := r.ParseMultipartForm(32 << 20)
	checkError(w, err)

	files := r.MultipartForm.File["file"][0]
	fileName = files.Filename

	segments := strings.Split(fileName, ".")
	extension := segments[len(segments)-1]
	if extension != "csv" {
		respondWithError(w, http.StatusBadRequest, "Please provide excel file having csv extension")
		return
	}
	fileBytes := ReadCSV(files)

	respondWithJson(w, http.StatusAccepted, fileBytes)
}

// ReadCSV File
func ReadCSV(file *multipart.FileHeader) string {
	csvFile, err := file.Open()

	if err != nil {
		log.Fatal("The file is not found || wrong root")
	}
	defer csvFile.Close()

	reader := csv.NewReader(csvFile)
	content, _ := reader.ReadAll()

	if len(content) < 1 {
		log.Fatal("Something wrong, the file maybe empty or length of the lines are not the same")
	}

	headersArr := make([]string, 0)
	for _, headE := range content[0] {
		headersArr = append(headersArr, headE)
	}

	//Remove the header row
	content = content[1:]

	var buffer bytes.Buffer
	buffer.WriteString("[")
	for i, d := range content {
		buffer.WriteString("{")
		for j, y := range d {
			buffer.WriteString(`"` + headersArr[j] + `":`)
			buffer.WriteString((`"` + y + `"`))
			//end of property
			if j < len(d)-1 {
				buffer.WriteString(",")
			}

		}
		//end of object of the array
		buffer.WriteString("}")
		if i < len(content)-1 {
			buffer.WriteString(",")
		}
	}

	buffer.WriteString(`]`)
	//rawMessage := json.RawMessage(buffer.String())
	//x, err := json.MarshalIndent(rawMessage, "", "  ")
	f := buffer.String()
	f = strings.ReplaceAll(f, `\`, "")
	fmt.Println(f)
	return f
}

func fetchSheetName(fileName string) (string, error) {
	var name string
	xlsx, err := excelize.OpenFile(dir + fileName)
	if err != nil {
		log.Println(err)
		return name, err
	}
	i := 0
	for _, sheet := range xlsx.GetSheetMap() {
		if i == 0 {
			name = sheet
		}
		i++
	}
	return name, err
}

func conversion2(w http.ResponseWriter, file string) {
	var (
		headers []string
		result  []*map[string]interface{}
		wb      = new(excelize.File)
		err     error
	)
	wb, err = excelize.OpenFile(dir + file)
	if err != nil {
		respondWithError(w, http.StatusBadRequest, "Error while open file")
		return
	}
	// Get all the rows in the Sheet.
	sheetName, _ := fetchSheetName(file)
	rows, _ := wb.GetRows(sheetName)
	headers = rows[0]
	for _, row := range rows[1:] {
		var tmpMap = make(map[string]interface{})
		for j, v := range row {
			tmpMap[strings.Join(strings.Split(headers[j], " "), "")] = v
		}
		result = append(result, &tmpMap)
	}
	respondWithJson(w, http.StatusAccepted, result)
}

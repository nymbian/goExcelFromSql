package main

import (
	"database/sql"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	_ "github.com/go-sql-driver/mysql"
)

//load conf
func loadConf(filename string) *map[string]string {
	bytes, err := ioutil.ReadFile(filename)
	checkErr(err)

	conf := make(map[string]string)

	json.Unmarshal(bytes, &conf)
	return &conf
}

func main() {
	start := time.Now()
	conf := *loadConf("conf.json")

	dsn := conf["dsn"]
	// Open database connection
	db, err := sql.Open("mysql", dsn)
	checkErr(err)
	defer db.Close()

	query := conf["query"]

	resultPointer, columnsPointer := sqlFetch(db, query)

	excel(resultPointer, columnsPointer)
	end := time.Now()
	fmt.Println("total time : ", timeFriendly(end.Sub(start).Seconds()))
}

//get result and columns
func sqlFetch(db *sql.DB, query string) (*[]map[string]string, *[]string) {

	// Execute the query
	rows, err := db.Query(query)
	checkErr(err)
	// Get column names
	columns, err := rows.Columns()
	checkErr(err)
	// Make a slice for the values
	values := make([]sql.RawBytes, len(columns))
	// rows.Scan wants '[]interface{}' as an argument, so we must copy the
	// references into such a slice
	// See http://code.google.com/p/go-wiki/wiki/InterfaceSlice for details
	scanArgs := make([]interface{}, len(values))
	for i := range values {
		scanArgs[i] = &values[i]
	}
	result := make([]map[string]string, 0)
	// Fetch rows
	for rows.Next() {
		// get RawBytes from data
		err = rows.Scan(scanArgs...)
		checkErr(err)
		// Now do something with the data.
		// Here we just print each column as a string.
		var value string
		vmap := make(map[string]string, len(scanArgs))
		for i, col := range values {
			// Here we can check if the value is nil (NULL value)
			if col == nil {
				value = "NULL"
			} else {
				value = string(col)
			}
			vmap[columns[i]] = value
			//fmt.Println(columns[i], ": ", value)
		}
		result = append(result, vmap)
		//fmt.Println("-----------------------------------")
	}
	if err = rows.Err(); err != nil {
		panic(err.Error()) // proper error handling instead of panic in your app
	}
	return &result, &columns

}

func excel(resultPointer *[]map[string]string, columnsPointer *[]string) {
	//xlsx := excelize.CreateFile()
	xlsx := excelize.NewFile()

	//Set value of a cell.
	result := *resultPointer
	columns := *columnsPointer

	//fmt.Println(result)
	fmt.Println(columns)
	//categories := map[string]string{"A1": "Small", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	//values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}

	categories := make(map[string]string, 0)

	for k, v := range columns {
		key := precessCategories(k)
		categories[key+"1"] = v
	}
	//fmt.Println(categories)
	for k, v := range categories {
		xlsx.SetCellValue("Sheet1", k, v)
	}

	values := make(map[string]string, 0)

	for k1, v1 := range result {

		//fmt.Println(v1)
		c := 0
		for k2, v2 := range v1 {

			i := getArrKey(columns, k2)
			key := precessCategories(i) + strconv.Itoa(k1+2)
			values[key] = v2
			//fmt.Println(key)

			c++
		}

	}
	//fmt.Println(values)
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}

	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(2)
	// Save xlsx file by the given path.
	err := xlsx.SaveAs("workbook.xlsx")
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

//map is not ordered
func getArrKey(arr []string, value string) int {
	for k, v := range arr {
		if v == value {
			return k
		}
	}
	return -1
}

//excel
func precessCategories(k int) string {
	az := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if k < 26 {
		return string(az[k])
	} else {
		k1 := int((k + 1) / 26)
		k2 := (k + 1) % 26
		return string(az[k1]) + string(az[k2])
	}
}

func checkErr(err error) {
	if err != nil {
		panic(err)
	}
}

// time format fridnely
func timeFriendly(second float64) string {

	if second < 1 {
		return strconv.Itoa(int(second*1000)) + "毫秒"
	} else if second < 60 {
		return strconv.Itoa(int(second)) + "秒" + timeFriendly(second-float64(int(second)))
	} else if second >= 60 && second < 3600 {
		return strconv.Itoa(int(second/60)) + "分" + timeFriendly(second-float64(int(second/60)*60))
	} else if second >= 3600 && second < 3600*24 {
		return strconv.Itoa(int(second/3600)) + "小时" + timeFriendly(second-float64(int(second/3600)*3600))
	} else if second > 3600*24 {
		return strconv.Itoa(int(second/(3600*24))) + "天" + timeFriendly(second-float64(int(second/(3600*24))*(3600*24)))
	}
	return ""
}

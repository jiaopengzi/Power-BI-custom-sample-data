// FilePath    : internal/data/data.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 内嵌地理 (省/市/区) 与姓名基础数据的加载.

// Package data 负责加载内嵌的地理 (省/市/区) 与姓名基础数据,
// 数据来源于原 VBA 的 AddressProvince/AddressCity/AddressDistrict/FirstName/LastName.
package data

import (
	"bufio"
	"embed"
	"strconv"
	"strings"
	"sync"

	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

//go:embed assets/province.csv assets/city.csv assets/district.csv assets/first_name.txt assets/last_name.txt
var assets embed.FS

// Province 省份记录, 对应 D01_省份表 数据源.
//   - RegionID, 大区 ID.
//   - ProvinceID, 省 ID.
//   - Full, 省全称.
//   - Short1, 省简称 1.
//   - Short2, 省简称 2.
//   - Lat, 纬度.
//   - Lng, 经度.
type Province struct {
	RegionID   int
	ProvinceID int
	Full       string
	Short1     string
	Short2     string
	Lat        float64
	Lng        float64
}

// City 城市记录, 对应 D02_城市表 数据源.
type City struct {
	ProvinceID int
	CityID     int
	Name       string
	Lat        float64
	Lng        float64
}

// District 区县记录, 对应 D03_区县表 数据源.
type District struct {
	CityID     int
	DistrictID int
	Name       string
	Lat        float64
	Lng        float64
}

// Region 大区记录, 对应 D00_大区表 (原 VBA 硬编码).
type Region struct {
	RegionID int
	Name     string
	Manager  string
	CityID   int
	CityName string
	Lat      float64
	Lng      float64
}

// Dataset 聚合全部基础数据及常用索引.
type Dataset struct {
	Regions    []Region
	Provinces  []Province
	Cities     []City
	Districts  []District
	FirstNames []string
	LastNames  []string

	// DistrictToProvince 区县 ID -> 省 ID 的索引, 用于销售目标按省汇总.
	DistrictToProvince map[int]int
	// ProvinceByID 省 ID -> 省份记录.
	ProvinceByID map[int]Province
}

var (
	loadOnce sync.Once
	dataset  *Dataset
	loadErr  error
)

// Load 加载并缓存全部内嵌基础数据 (线程安全, 仅解析一次).
// 返回值 *Dataset, 数据集; error, 解析失败时非 nil.
func Load() (*Dataset, error) {
	loadOnce.Do(func() {
		dataset, loadErr = parse()
	})
	return dataset, loadErr
}

// hardcodedRegions 原 VBA DataTableD0 中硬编码的六大区.
var hardcodedRegions = []Region{
	{1, "东区", "欧阳经往", 310000, "上海", 31.231518, 121.471518},
	{2, "西区", "焦阿灰", 510100, "成都", 30.659518, 104.065518},
	{3, "南区", "左丘垂漫", 440100, "广州", 23.125518, 113.280518},
	{4, "北区", "左烈佐", 210100, "沈阳", 41.796518, 123.429518},
	{5, "中区", "焦仔耘", 110000, "北京", 39.901518, 116.401518},
	{6, "港澳台", "安修谊", 810000, "香港", 22.320518, 114.173518},
}

// parse 解析全部内嵌资源为 Dataset.
// 返回值 *Dataset, 数据集; error, 出错时非 nil.
func parse() (*Dataset, error) {
	ds := &Dataset{
		Regions:            hardcodedRegions,
		DistrictToProvince: make(map[int]int),
		ProvinceByID:       make(map[int]Province),
	}

	provRows, err := readCSV("assets/province.csv")
	if err != nil {
		return nil, err
	}
	for _, r := range provRows {
		p := Province{
			RegionID:   atoi(r[0]),
			ProvinceID: atoi(r[1]),
			Full:       r[2],
			Short1:     r[3],
			Short2:     r[4],
			Lat:        atof(r[5]),
			Lng:        atof(r[6]),
		}
		ds.Provinces = append(ds.Provinces, p)
		ds.ProvinceByID[p.ProvinceID] = p
	}

	cityRows, err := readCSV("assets/city.csv")
	if err != nil {
		return nil, err
	}
	cityToProvince := make(map[int]int, len(cityRows))
	for _, r := range cityRows {
		c := City{
			ProvinceID: atoi(r[0]),
			CityID:     atoi(r[1]),
			Name:       r[2],
			Lat:        atof(r[3]),
			Lng:        atof(r[4]),
		}
		ds.Cities = append(ds.Cities, c)
		cityToProvince[c.CityID] = c.ProvinceID
	}

	distRows, err := readCSV("assets/district.csv")
	if err != nil {
		return nil, err
	}
	for _, r := range distRows {
		d := District{
			CityID:     atoi(r[0]),
			DistrictID: atoi(r[1]),
			Name:       r[2],
			Lat:        atof(r[3]),
			Lng:        atof(r[4]),
		}
		ds.Districts = append(ds.Districts, d)
		if pid, ok := cityToProvince[d.CityID]; ok {
			ds.DistrictToProvince[d.DistrictID] = pid
		}
	}

	if ds.FirstNames, err = readNames("assets/first_name.txt"); err != nil {
		return nil, err
	}
	if ds.LastNames, err = readNames("assets/last_name.txt"); err != nil {
		return nil, err
	}
	return ds, nil
}

// readCSV 读取内嵌 CSV, 跳过表头, 返回字段切片列表.
//   - name, 内嵌资源路径.
//
// 返回值 [][]string, 数据行; error, 出错时非 nil.
func readCSV(name string) ([][]string, error) {
	f, err := assets.Open(name)
	if err != nil {
		return nil, err
	}
	defer util.CloseQuietly(f)

	var rows [][]string
	sc := bufio.NewScanner(f)
	sc.Buffer(make([]byte, 0, 64*1024), 1024*1024)
	first := true
	for sc.Scan() {
		line := strings.TrimSpace(sc.Text())
		if first {
			first = false
			continue
		}
		if line == "" {
			continue
		}
		rows = append(rows, strings.Split(line, ","))
	}
	return rows, sc.Err()
}

// readNames 读取逗号分隔的姓名文件为字符串切片.
//   - name, 内嵌资源路径.
//
// 返回值 []string, 姓名列表; error, 出错时非 nil.
func readNames(name string) ([]string, error) {
	b, err := assets.ReadFile(name)
	if err != nil {
		return nil, err
	}
	parts := strings.Split(strings.TrimSpace(string(b)), ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		if p = strings.TrimSpace(p); p != "" {
			out = append(out, p)
		}
	}
	return out, nil
}

// atoi 宽松地将字符串转为 int, 失败返回 0.
func atoi(s string) int {
	n, _ := strconv.Atoi(strings.TrimSpace(s)) //nolint:errcheck // 宽松解析, 失败取 0
	return n
}

// atof 宽松地将字符串转为 float64, 失败返回 0.
func atof(s string) float64 {
	f, _ := strconv.ParseFloat(strings.TrimSpace(s), 64) //nolint:errcheck // 宽松解析, 失败取 0
	return f
}

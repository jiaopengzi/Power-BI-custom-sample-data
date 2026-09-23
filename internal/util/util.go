// FilePath    : internal/util/util.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : VBA 移植所需的数值/随机/格式化辅助函数.

// Package util 提供从 VBA 移植所需的数值/随机/格式化辅助函数.
// 目标是保持业务逻辑一致 (非逐字节复现), 因此随机数使用 Go 标准库,
// 但 Round 采用与 VBA 一致的银行家舍入 (round half to even).
package util

import (
	"io"
	"math"
	"math/rand"
	"strconv"
	"time"
)

// SilentClose 执行清理函数并忽略其错误, 仅用于错误处理/延迟释放路径.
//   - fn, 返回 error 的清理函数 (如 io.Closer.Close).
func SilentClose(fn func() error) {
	_ = fn() //nolint:errcheck // 清理路径故意忽略关闭错误
}

// CloseQuietly 关闭一个 io.Closer 并忽略错误, 仅用于延迟释放场景.
//   - c, 待关闭的资源.
func CloseQuietly(c io.Closer) {
	_ = c.Close() //nolint:errcheck // 清理路径故意忽略关闭错误
}

// Rand 封装随机源, 对应 VBA 的 Rnd() 与 Randomize.
type Rand struct {
	r *rand.Rand
}

// NewRand 创建一个以当前时间为种子的随机源.
// 返回值 *Rand, 随机源实例.
func NewRand() *Rand {
	// #nosec G404 示例数据生成无需加密级随机, 使用 math/rand 即可.
	return &Rand{r: rand.New(rand.NewSource(time.Now().UnixNano()))}
}

// NewRandSeed 创建一个指定种子的随机源, 便于测试复现.
//   - seed, 随机种子.
//
// 返回值 *Rand, 随机源实例.
func NewRandSeed(seed int64) *Rand {
	// #nosec G404 示例数据生成无需加密级随机, 使用 math/rand 即可.
	return &Rand{r: rand.New(rand.NewSource(seed))}
}

// F 返回 [0, 1) 的随机浮点数, 对应 VBA 的 Rnd().
// 返回值 float64, 随机数.
func (rd *Rand) F() float64 {
	return rd.r.Float64()
}

// Norm 返回标准正态分布 N(0, 1) 的随机抽样 (Box-Muller 变换),
// 供需要符合客观规律的分布 (如对数正态价格) 使用.
// 返回值 float64, 标准正态随机数.
func (rd *Rand) Norm() float64 {
	u1 := rd.r.Float64()
	u2 := rd.r.Float64()
	if u1 <= 0 {
		u1 = 1e-12 // 避免 log(0).
	}
	return math.Sqrt(-2*math.Log(u1)) * math.Cos(2*math.Pi*u2) // #nosec G115 math 库函数, 无转换溢出
}

// RoundBankers 银行家舍入 (round half to even), 对应 VBA 的 Round.
//   - x, 待舍入的数值.
//   - places, 保留的小数位数.
//
// 返回值 float64, 舍入结果.
func RoundBankers(x float64, places int) float64 {
	if places < 0 {
		places = 0
	}
	pow := math.Pow(10, float64(places))
	return math.RoundToEven(x*pow) / pow
}

// RoundInt 对数值做 0 位银行家舍入并转为 int, 对应 VBA 的 Round(x, 0).
//   - x, 待舍入的数值.
//
// 返回值 int, 舍入后的整数.
func RoundInt(x float64) int {
	return int(math.RoundToEven(x))
}

// PadInt 将整数格式化为固定宽度的零填充字符串, 对应 VBA 的 Format(n, "000...").
//   - n, 待格式化的整数.
//   - width, 目标宽度.
//
// 返回值 string, 零填充后的字符串.
func PadInt(n, width int) string {
	s := strconv.Itoa(n)
	for len(s) < width {
		s = "0" + s
	}
	return s
}

// Letter 返回从 'A' 起偏移 n 个位置的大写字母, 用于产品分类/门店名生成.
//   - n, 偏移量, 取值受限于 0-25 的小范围.
//
// 返回值 string, 对应字母.
func Letter(n int) string {
	return string(rune('A' + n)) // #nosec G115 n 取值受限于小范围, 不会溢出
}

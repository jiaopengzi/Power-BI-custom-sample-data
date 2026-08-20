// FilePath    : internal/csvw/writer.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 带 UTF-8 BOM 的 CSV 写入器.

// Package csvw 提供带 UTF-8 BOM 的 CSV 写入器, 便于 Excel 直接打开中文不乱码.
package csvw

import (
	"bufio"
	"encoding/csv"
	"os"
	"path/filepath"

	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// utf8BOM UTF-8 字节顺序标记.
var utf8BOM = []byte{0xEF, 0xBB, 0xBF}

// Writer 封装一个面向单个 CSV 文件的写入器.
type Writer struct {
	f   *os.File
	buf *bufio.Writer
	w   *csv.Writer
}

// Create 创建 (或覆盖) 一个 CSV 文件并写入 BOM 与表头.
//   - dir, 目标目录.
//   - fileName, 文件名 (含 .csv).
//   - header, CSV 表头字段.
//
// 返回值 *Writer, 写入器; error, 出错时非 nil.
func Create(dir, fileName string, header []string) (*Writer, error) {
	if err := os.MkdirAll(dir, 0o750); err != nil {
		return nil, err
	}
	f, err := os.Create(filepath.Join(dir, fileName)) // #nosec G304 目录由上层校验
	if err != nil {
		return nil, err
	}
	buf := bufio.NewWriterSize(f, 256*1024)
	if _, err = buf.Write(utf8BOM); err != nil {
		util.CloseQuietly(f)
		return nil, err
	}
	w := csv.NewWriter(buf)
	if err = w.Write(header); err != nil {
		util.CloseQuietly(f)
		return nil, err
	}
	return &Writer{f: f, buf: buf, w: w}, nil
}

// OpenAppend 以追加模式打开一个已存在的 CSV 文件 (用于增量更新, 不写表头).
//   - dir, 目标目录.
//   - fileName, 文件名.
//
// 返回值 *Writer, 写入器; error, 出错时非 nil.
func OpenAppend(dir, fileName string) (*Writer, error) {
	f, err := os.OpenFile(filepath.Join(dir, fileName), os.O_APPEND|os.O_WRONLY, 0o600) // #nosec G304
	if err != nil {
		return nil, err
	}
	buf := bufio.NewWriterSize(f, 256*1024)
	return &Writer{f: f, buf: buf, w: csv.NewWriter(buf)}, nil
}

// Write 写入一行记录.
//   - record, 一行字段.
//
// 返回值 error, 出错时非 nil.
func (w *Writer) Write(record []string) error {
	return w.w.Write(record)
}

// Close 刷新缓冲并关闭文件.
// 返回值 error, 出错时非 nil.
func (w *Writer) Close() error {
	w.w.Flush()
	if err := w.w.Error(); err != nil {
		util.CloseQuietly(w.f)
		return err
	}
	if err := w.buf.Flush(); err != nil {
		util.CloseQuietly(w.f)
		return err
	}
	return w.f.Close()
}

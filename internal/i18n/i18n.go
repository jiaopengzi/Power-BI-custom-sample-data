// FilePath    : internal/i18n/i18n.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 运行时可切换语言的文案管理器.

// Package i18n 提供运行时可切换的中英文文案查询.
// 语言标识与 config.Locale 对齐 (zh-cn 默认 / en-us), 文案表见 messages.go.
package i18n

import (
	"strings"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
)

// Manager 持有当前语言并提供文案查询, 供界面在运行时切换语言.
type Manager struct {
	locale config.Locale
}

// NewManager 创建文案管理器.
//   - loc, 初始语言, 非 en-us 一律回退为 zh-cn.
//
// 返回值 *Manager, 管理器实例.
func NewManager(loc config.Locale) *Manager {
	m := &Manager{}
	m.SetLocale(loc)
	return m
}

// Locale 返回当前语言.
// 返回值 config.Locale, 当前语言.
func (m *Manager) Locale() config.Locale {
	return m.locale
}

// SetLocale 设置当前语言, 非 en-us 一律回退为 zh-cn.
//   - loc, 目标语言.
func (m *Manager) SetLocale(loc config.Locale) {
	if loc == config.LocaleEnUS {
		m.locale = config.LocaleEnUS
		return
	}
	m.locale = config.LocaleZhCN
}

// T 返回指定键的当前语言文案, 缺失时回退为键名本身.
//   - key, 点号分隔的文案键.
//
// 返回值 string, 文案.
func (m *Manager) T(key string) string {
	e, ok := catalog[key]
	if !ok {
		return key
	}
	if m.locale == config.LocaleEnUS {
		return e.en
	}
	return e.zh
}

// Tf 返回指定键的文案并以 {name} 占位符替换 vars 中的键值.
//   - key, 文案键.
//   - vars, 占位符名称到替换值的映射 (名称不含花括号).
//
// 返回值 string, 替换后的文案.
func (m *Manager) Tf(key string, vars map[string]string) string {
	s := m.T(key)
	if len(vars) == 0 {
		return s
	}
	pairs := make([]string, 0, len(vars)*2)
	for k, v := range vars {
		pairs = append(pairs, "{"+k+"}", v)
	}
	return strings.NewReplacer(pairs...).Replace(s)
}

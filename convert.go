package xlstruct

import (
	"errors"
	"log"
	"reflect"

	"github.com/tealeg/xlsx"
)

var (
	ErrNotPointer = errors.New("supports only unmarshaling to pointer")
	ErrNotSlice   = errors.New("supports only unmarshaling to slice")
	ErrNotStruct  = errors.New("supports only unmarshaling to slice that element is struct")
)

type unmarshaler struct {
	isPtr    bool
	typ      reflect.Type
	sch      map[string]int
	cellSize int
	idxm     map[string]int
	typm     map[string]reflect.Type
}

func newUnmarshaler(v interface{}, headers []*xlsx.Cell) (*unmarshaler, error) {
	schema := map[string]int{}
	for i, cell := range headers {
		if cell.Value == "" {
			continue
		}
		schema[cell.Value] = i
	}
	t := reflect.TypeOf(v)
	if t.Kind() != reflect.Ptr {
		return nil, ErrNotPointer
	}
	t = t.Elem()

	if t.Kind() != reflect.Slice {
		return nil, ErrNotSlice
	}
	t = t.Elem()

	isPtr := false
	if t.Kind() == reflect.Ptr {
		isPtr = true
		t = t.Elem()
	}

	if t.Kind() != reflect.Struct {
		return nil, ErrNotStruct
	}

	// get FieldName => row[i] mapping
	idxm, typm := getFieldCellMaps(t, schema)
	return &unmarshaler{
		isPtr:    isPtr,
		typ:      t,
		sch:      schema,
		cellSize: len(headers),
		idxm:     idxm,
		typm:     typm,
	}, nil
}

// Unmarshal supports only struct pointer type
func Unmarshal(v interface{}, sheet *xlsx.Sheet, schidx int) error {
	h := sheet.Rows[schidx]
	um, err := newUnmarshaler(v, h.Cells)
	if err != nil {
		return err
	}

	// prepare results
	results := reflect.ValueOf(v).Elem() // ptr is not Settable

Loop:
	for i, row := range sheet.Rows {
		if len(row.Cells) < um.cellSize {
			break
		}
		if i == schidx {
			continue // skip schema line
		}
		n := reflect.New(um.typ).Elem()
		// assign key
		for k, idx := range um.idxm {
			field := n.FieldByName(k)
			cell := row.Cells[idx]
			val, err := getValue(um.typm[k], cell)
			if err != nil {
				log.Printf("bad value(%v) for field(%v) as type(%s)\n", cell.Value, k, field.Type())
				continue Loop
			}
			field.Set(val)
		}
		if um.isPtr {
			n = n.Addr()
		}
		results.Set(reflect.Append(results, n))
	}

	return nil
}

func getValue(typ reflect.Type, cell *xlsx.Cell) (reflect.Value, error) {
	var v interface{}
	var err error
	switch typ.Kind() {
	case reflect.Int:
		v, err = cell.Int()
	case reflect.Int64:
		v, err = cell.Int64()
	case reflect.Float64:
		v, err = cell.Float()
	case reflect.String:
		v = cell.String()
	}
	if err != nil {
		return reflect.Value{}, err
	}
	return reflect.ValueOf(v), nil
}

func getFieldCellMaps(t reflect.Type, sch map[string]int) (map[string]int, map[string]reflect.Type) {
	idxm := map[string]int{}
	typm := map[string]reflect.Type{}
	for i := 0; i < t.NumField(); i++ {
		// skip the fields that don't have excel tag
		field := t.Field(i)
		tag := field.Tag.Get("excel")
		if tag == "" {
			continue
		}

		idx, ok := sch[tag]
		if !ok {
			continue
		}
		fm := field.Name
		idxm[fm] = idx
		typm[fm] = field.Type
	}
	return idxm, typm
}

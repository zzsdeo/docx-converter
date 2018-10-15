// Copyright 2017 Baliance. All rights reserved.
//
// DO NOT EDIT: generated by gooxml ECMA-376 generator
//
// Use of this source code is governed by the terms of the Affero GNU General
// Public License version 3.0 as published by the Free Software Foundation and
// appearing in the file LICENSE included in the packaging of this file. A
// commercial license can be purchased by contacting sales@baliance.com.

package wml

import (
	"encoding/xml"

	"baliance.com/gooxml"
)

type EG_CellMarkupElements struct {
	// Table Cell Insertion
	CellIns *CT_TrackChange
	// Table Cell Deletion
	CellDel *CT_TrackChange
	// Vertically Merged/Split Table Cells
	CellMerge *CT_CellMergeTrackChange
}

func NewEG_CellMarkupElements() *EG_CellMarkupElements {
	ret := &EG_CellMarkupElements{}
	return ret
}

func (m *EG_CellMarkupElements) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if m.CellIns != nil {
		secellIns := xml.StartElement{Name: xml.Name{Local: "w:cellIns"}}
		e.EncodeElement(m.CellIns, secellIns)
	}
	if m.CellDel != nil {
		secellDel := xml.StartElement{Name: xml.Name{Local: "w:cellDel"}}
		e.EncodeElement(m.CellDel, secellDel)
	}
	if m.CellMerge != nil {
		secellMerge := xml.StartElement{Name: xml.Name{Local: "w:cellMerge"}}
		e.EncodeElement(m.CellMerge, secellMerge)
	}
	return nil
}

func (m *EG_CellMarkupElements) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	// initialize to default
lEG_CellMarkupElements:
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name {
			case xml.Name{Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main", Local: "cellIns"}:
				m.CellIns = NewCT_TrackChange()
				if err := d.DecodeElement(m.CellIns, &el); err != nil {
					return err
				}
			case xml.Name{Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main", Local: "cellDel"}:
				m.CellDel = NewCT_TrackChange()
				if err := d.DecodeElement(m.CellDel, &el); err != nil {
					return err
				}
			case xml.Name{Space: "http://schemas.openxmlformats.org/wordprocessingml/2006/main", Local: "cellMerge"}:
				m.CellMerge = NewCT_CellMergeTrackChange()
				if err := d.DecodeElement(m.CellMerge, &el); err != nil {
					return err
				}
			default:
				gooxml.Log("skipping unsupported element on EG_CellMarkupElements %v", el.Name)
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			break lEG_CellMarkupElements
		case xml.CharData:
		}
	}
	return nil
}

// Validate validates the EG_CellMarkupElements and its children
func (m *EG_CellMarkupElements) Validate() error {
	return m.ValidateWithPath("EG_CellMarkupElements")
}

// ValidateWithPath validates the EG_CellMarkupElements and its children, prefixing error messages with path
func (m *EG_CellMarkupElements) ValidateWithPath(path string) error {
	if m.CellIns != nil {
		if err := m.CellIns.ValidateWithPath(path + "/CellIns"); err != nil {
			return err
		}
	}
	if m.CellDel != nil {
		if err := m.CellDel.ValidateWithPath(path + "/CellDel"); err != nil {
			return err
		}
	}
	if m.CellMerge != nil {
		if err := m.CellMerge.ValidateWithPath(path + "/CellMerge"); err != nil {
			return err
		}
	}
	return nil
}

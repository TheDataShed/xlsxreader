package xlsxreader

import "encoding/xml"

func getCharData(d *xml.Decoder) (string, error) {
	tok, err := d.RawToken()
	if err != nil {
		return "", err
	}

	cdata, ok := tok.(xml.CharData)
	if !ok {
		return "", xml.UnmarshalError("unexpected token where chardata expected")
	}

	return string(cdata), nil
}

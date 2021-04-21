package xlsxreader

import "encoding/xml"

func getCharData(d *xml.Decoder) (string, error) {
	rawToken, err := d.RawToken()
	if err != nil {
		return "", err
	}

	cdata, ok := rawToken.(xml.CharData)
	if !ok {
		return "", xml.UnmarshalError("expected chardata to be present, but none was found")
	}

	return string(cdata), nil
}

from iphone_path.logic import IPhonePathImpl, IPHONE_SCHEME, decode_parts, encode_component


def test_iphone_path_scheme_overrides_builtin():
    assert IPHONE_SCHEME["class"] is IPhonePathImpl
    assert IPHONE_SCHEME["override"] is True


def test_uri_parts_are_decoded_without_device_access():
    assert decode_parts("photos%20and%20videos/DCIM") == ["photos and videos", "DCIM"]


def test_child_names_are_encoded_as_uri_components():
    assert encode_component("My Photos") == "My%20Photos"

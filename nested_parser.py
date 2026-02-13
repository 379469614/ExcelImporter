#encoding=utf-8
import string

_OPEN_TO_CLOSE = {
  '{': '}',
  '[': ']',
  '(': ')',
}

_CLOSE_SET = set(_OPEN_TO_CLOSE.values())


def split_top_level(text, delimiter, skip_empty = False):
  if text is None:
    return []
  if not delimiter:
    raise ValueError('delimiter can not be empty')

  values = []
  stack = []
  start = 0
  i = 0
  n = len(text)

  while i < n:
    c = text[i]

    if c == '\\' and i + 1 < n:
      i += 2
      continue

    if c in _OPEN_TO_CLOSE:
      stack.append(_OPEN_TO_CLOSE[c])
      i += 1
      continue

    if c in _CLOSE_SET:
      if not stack or c != stack[-1]:
        raise ValueError('%s is not a legal nested expression' % text)
      stack.pop()
      i += 1
      continue

    if not stack and text.startswith(delimiter, i):
      value = text[start:i]
      if not skip_empty or value:
        values.append(value)
      i += len(delimiter)
      start = i
      continue

    i += 1

  if stack:
    raise ValueError('%s is not a legal nested expression' % text)

  value = text[start:]
  if not skip_empty or value:
    values.append(value)

  return values


def unwrap_container(text, begin, end):
  value = text.strip()
  if len(value) >= 2 and value[0] == begin and value[-1] == end:
    return value[1:-1]
  return value


def split_list_values(value):
  return split_top_level(unwrap_container(value, '[', ']'), ',', True)


def split_obj_type_fields(type_, separator):
  return split_top_level(unwrap_container(type_, '{', '}'), separator, True)


def split_obj_values(value, separator):
  return split_top_level(unwrap_container(value, '{', '}'), separator, True)


def split_field_declaration(text):
  declaration = text.strip()
  if not declaration:
    raise ValueError('field declaration can not be empty')

  values = []
  stack = []
  start = None
  i = 0
  n = len(declaration)

  while i < n:
    c = declaration[i]

    if c == '\\' and i + 1 < n:
      if start is None:
        start = i
      i += 2
      continue

    if c in _OPEN_TO_CLOSE:
      if start is None:
        start = i
      stack.append(_OPEN_TO_CLOSE[c])
      i += 1
      continue

    if c in _CLOSE_SET:
      if not stack or c != stack[-1]:
        raise ValueError('%s is not a legal field declaration' % declaration)
      stack.pop()
      i += 1
      continue

    if c in string.whitespace and not stack:
      if start is not None:
        values.append(declaration[start:i])
        start = None
      i += 1
      continue

    if start is None:
      start = i
    i += 1

  if stack:
    raise ValueError('%s is not a legal field declaration' % declaration)

  if start is not None:
    values.append(declaration[start:])

  if len(values) < 2:
    raise ValueError('%s is not a legal field declaration' % declaration)

  return (' '.join(values[:-1]), values[-1])
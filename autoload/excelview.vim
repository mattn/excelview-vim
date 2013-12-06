function! s:loadSharedStrings(f)
  let ss = []
  try
    let xml = system(printf("unzip -p -- %s xl/sharedStrings.xml", shellescape(a:f)))
    let doc = webapi#xml#parse(xml)
    for si in doc.childNodes("si")
      let t = si.childNode("t")
      call add(ss, t.value())
    endfor
  catch
  endtry
  return ss
endfunction

function! s:loadSheetData(f, s)
  let xml = system(printf("unzip -p -- %s xl/worksheets/sheet%d.xml", shellescape(a:f), a:s))

  let ss = s:loadSharedStrings(a:f)
  let doc = webapi#xml#parse(xml)
  let rows = doc.childNode("sheetData").childNodes("row")
  let cells = map(range(1, 256), 'map(range(1,256), "''''")')
  let aa = char2nr('A')
  for row in rows
    for col in row.childNodes("c")
      let r = col.attr["r"]
      let nv = col.childNode("v")
      let v = empty(nv) ? "" : nv.value()
      if has_key(col.attr, "s") && col.attr["s"] == "2"
        let v = strftime("%Y/%m/%d %H:%M:%S", (v - 25569) * 86400 - 32400)
      endif
      if has_key(col.attr, "t") && col.attr["t"] == "s"
        let v = ss[v]
      endif
      let x = char2nr(r[0]) - aa
      let y = matchstr(r, '\d\+')
      let cells[y][x+1] = v
    endfor
  endfor
  for y in range(len(cells)-1)
    let cells[y+1][0] = y + 1
  endfor
  for x in range(len(cells[0])-1)
    let nx = x / 26
    if nx == 0
      let cells[0][x+1] = nr2char(aa+x)
    else
      let cells[0][x+1] = nr2char(aa+nx-1) . nr2char(aa+x%26)
    endif
  endfor
  return cells
endfunction

function! s:fillColumns(rows)
  let rows = a:rows
  if type(rows) != 3 || type(rows[0]) != 3
    return [[]]
  endif
  let cols = len(rows[0])
  for c in range(cols)
    let m = 0
    let w = range(len(rows))
    for r in range(len(w))
      if type(rows[r][c]) == 2
        let s = string(rows[r][c])
      endif
      let w[r] = strdisplaywidth(rows[r][c])
      let m = max([m, w[r]])
    endfor
    for r in range(len(w))
      let rows[r][c] = ' ' . rows[r][c] . repeat(' ', m - w[r]) . ' '
    endfor
  endfor
  return rows
endfunction

function! excelview#view(...) abort
  if a:0 > 2
    echohl Error | echon "Usage: :ExcelView [filename] {[sheet-number]}" | echohl None
	return
  endif
  let [f, s] = a:0 == 1 ? [a:1, 1] : [a:1, a:2]
  try
    let data = s:loadSheetData(f, s)
  catch
    let e = v:exception
    echohl Error | echon printf("Error while loading sheet%d: %s", s, e) | echohl None
	return
  endtry
  new
  setlocal noswapfile buftype=nofile bufhidden=delete nowrap norightleft modifiable nolist nonumber
  let data = s:fillColumns(data) 
  let sep = "+" . join(map(copy(data[0]), 'repeat("-", len(v:val))'), '+') . "+"
  call setline(1, sep)
  let r = 2
  for row in data
    let line = join(row, '|')
    call setline(r, '|'.join(row, '|').'|')
    call setline(r + 1, sep)
    let r += 2
  endfor
  setlocal nomodifiable
endfunction

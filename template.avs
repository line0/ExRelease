vidfile = "{{?VIDEO_FILE}}"
subfile = "{{?SUBTITLE_FILE}}"

c1  = FFVideoSource (vidfile, enable10bithack=false)
c2  = FFVideoSource (vidfile, enable10bithack=true, varprefix="c1_")
w   = c1.Width ()
h   = c1.Height ()
wm4 = (w + 3) / 4 * 4
hm4 = (h + 3) / 4 * 4
b10 = (c2.Height () > c1.Height ())

target_width = {{?TARGET_RES_X}}
target_height = {{?TARGET_RES_Y}}
display_timecode = false
display_subtitles = true
group = "{{?GROUP_NAME}}"

matrix   = (c1_FFCOLOR_SPACE == 1) ? "709"
\        : (c1_FFCOLOR_SPACE >= 5) ? "601" : Dither_undef ()
par      = (c1_FFSAR_NUM > 0 && c1_FFSAR_DEN > 0)
\        ? Float (c1_FFSAR_NUM) / Float (c1_FFSAR_DEN)
\        : Dither_undef ()
tv_range = (c1_FFCOLOR_RANGE != 2)

(b10) ? c2 : c1

(StrLen(subfile) > 0 && display_subtitles) ?
\                                    (b10) ? Dither_limit_dif16(DitherPost(mode=-1).TextSub(subfile).Dither_convert_8_to_16 (), thr=1.0, elast=2.0) : TextSub(subfile)
\: nop()

(VersionNumber () >= 2.6) ? last
\ : (b10) ? Dither_addborders16 (0, 0, wm4 - w, hm4 - h)
\         : AddBorders (0, 0, wm4 - w, hm4 - h)

(target_width != c1.Width() || target_height != c1.Height()) ? Dither_srgb_display (lsb_in=b10, matrix=matrix, par=par, tv_range=tv_range, w=target_width, h=target_height)
\ :                                                            Dither_srgb_display (lsb_in=b10, matrix=matrix, par=par, tv_range=tv_range)

cmt = (target_width > c1.Width() || target_height > c1.Height()) ? group + "\nupscaled" : group
FFInfoEx(vfrtime=display_timecode, cfrtime=false, frametype=false, x=last.width-6, y=4, cmt=group, align=9, shorttime=true)
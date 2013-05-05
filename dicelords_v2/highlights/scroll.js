var Tscroll_path_to_files = 'highlights/'

function Tscroll_init (id)
{
document.write ('<iframe id="Tscr' + id + '" scrolling=no frameborder=no src="' + Tscroll_path_to_files + 'scroll.html?' + id + '"></iframe>');
}
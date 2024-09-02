# light.tcl
package require Tk

# Define a light theme
ttk::style theme create light \
    -theme "default" \
    -background "#FFFFFF" \
    -foreground "#000000" \
    -fieldbackground "#F0F0F0" \
    -highlightcolor "#DDDDDD" \
    -highlightbackground "#DDDDDD" \
    -borderwidth 1 \
    -relief flat

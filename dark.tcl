# dark.tcl
package require Tk

# Define a dark theme
ttk::style theme create dark \
    -theme "default" \
    -background "#2E2E2E" \
    -foreground "#FFFFFF" \
    -fieldbackground "#3C3C3C" \
    -highlightcolor "#1E1E1E" \
    -highlightbackground "#1E1E1E" \
    -borderwidth 1 \
    -relief flat

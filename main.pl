use strict;
use warnings;
use Win32::GuiTest qw(FindWindowLike GetWindowText 
  SetForegroundWindow SendKeys);
 
$Win32::GuiTest::debug = 0; # Set to "1" to enable verbose mode
 
my @windows = FindWindowLike(0, "^Microsoft Excel", "^XLMAIN\$");
for (@windows) {
    print "$_>\t'", GetWindowText($_), "'\n";
    SetForegroundWindow($_);
    SendKeys("%fn~a{TAB}b{TAB}{BS}{DOWN}");
}


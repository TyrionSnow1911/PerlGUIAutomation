#create a perl script which slaves autoIT and does the follow
#open google chrome and check your email, if prompted for username and password enter them automatically



use strict;
use warnings;
use Win32;
use Win32::API;
use Win32::OLE;
my $Au3 = Win32::OLE->new("AutoItX3.Control");

my ($fname, $Curr, $fname1, $bbox, $bbox1, $file1, $error, $path, $filex, $filexx, $contentopf, $tocncx, $filess, $ebookmetadata1, $filecc, $bbox99, $snocounter);
my ($Line, $Variable, $PreMatch, $taglist, $taglist1, $mapfile, $mapbox, $maptext, $find, $replace);
my ($foottxt)  = ("");
my $fpath = $ARGV[0];
my $fpath2 = $ARGV[1];

#Win32::MsgBox("SET MATHTYPE TRANSLATIONS IF NOT!!!", MB_ICONSTOP, "SET MATHTYPE TRANSLATIONS IF NOT");
opendir (FF, $fpath) or die Win32::MsgBox("Directory not found!!!", MB_ICONSTOP, "DIRECTORY NOT FOUND");
my @contents3 = grep /(\.eps$)/, readdir FF;
foreach $file1 ( @contents3 )
    {
        #print "$file1 processing\n";
        $file1=~s/\.eps$//g;
        open($path, ">$fpath\\$file1\.txt");
        close($path);
        $Au3->run ('C:\Program Files\MathType\MathType.exe');
        $Au3->WinWaitActive('MathType - Untitled 1');
        $Au3->Sleep (500);
        $Au3->Send ('^o');
        $Au3->WinWaitActive('Open');
        $Au3->Sleep (1000);
        $Au3->send ("$fpath\\$file1\.eps");
        $Au3->Sleep (1500);
        $Au3->Send ('{ENTER}');
        $Au3->WinWaitActive('[CLASS:EQNWINCLASS]');
        $Au3->Sleep (1000);
        $Au3->Send ('^a');
        $Au3->Sleep (1000);
        $Au3->Send ('^c');
        $Au3->Sleep (1000);
        $Au3->run ('notepad.exe');
        $Au3->WinWaitActive('Untitled - Notepad');
        $Au3->Sleep (500);
        $Au3->Send ('^o');
        $Au3->WinWaitActive('Open');
        $Au3->send ("$fpath\\$file1\.txt");
        $Au3->Sleep (500);
        $Au3->Send ('{ENTER}');
        $Au3->WinWaitActive("$file1\.txt");
        $Au3->Send ('^v');
        $Au3->Sleep (1500);
        $Au3->Send ('!fx');
        $Au3->WinWaitActive('Notepad');
        $Au3->Send ('!y');
        $Au3->ProcessClose('mathtype.exe');                
        $Au3->Sleep (1000);
        print "$file1 completed\n";
    }

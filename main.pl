use strict;
use warnings;
use Win32;
use Win32::API;
use Win32::OLE;
use Cwd;
my $Au3 = Win32::OLE->new ("AutoItX3.Control");

my $rootDir = cwd;
my $testDir = "$rootDir\\Test_Cases";

#######################Generate addition test input file.###############################
open my $fileHandle, '>',
    "$testDir\\addition_test_input_file.txt" or die
        "Cannot create test case file: $testDir\\addition_test_input_file.txt!";
my $numOfTestCases = 100;

while ($numOfTestCases-- > 0) {
  my @numbers = genNumbers();
  addArrayelements($fileHandle,\@numbers);
}
close $fileHandle;
#######################################################################################


#######Generate text file of expected outputs from addition_test_input_file.txt########
open my $_fileHandle, '<',
    "$testDir\\addition_test_input_file.txt" or die
        "Cannot create test case file: $testDir\\addition_test_input_file.txt!";

open my $FH, '>', "$testDir\\expected_output.txt" or die
        "Cannot create test case file: $testDir\\expected_output.txt!";

while (my $line = <$_fileHandle>) {
chomp $line;
my @values = split /=/, $line;
my $n = scalar @values;
print $FH $values[$n-1];
print $FH "\n";
}
close $FH;
close $_fileHandle;
########################################################################################


#Write code to slave AutoIT with perl and execute each test case in test case input file.


#open calculator, add two numbers together then close after 2 seconds.
#$Au3->Run("calc.exe");
#$Au3->WinWaitActive("Calculator");
#$Au3->WinWaitActive("Calculator");
#$Au3->Send("{1}"); #1
#$Au3->Send("{+}"); # +
#$Au3->Send("{1}"); #1
#$Au3->Send("{ENTER}");
#$Au3->Sleep(2000);
#$Au3->WinClose("Calculator");







####Take each number in an array and add them together.Print to test case file in the form 1 + 2 + 3... = 100.####
sub addArrayelements {
  my $fh = shift;
  my $arr_ref = shift;
  my @arr = @$arr_ref;
  my $size = scalar @arr;
  my $sum = 0;

  for (my $i = 0; $i < $size; $i++) {
    print $fh $arr[$i];
    $sum += $arr[$i];
    if ($i == $size - 1) {
      print $fh "=";
    } else {
      print $fh "+";
    }
  }

  print $fh "$sum\n";
}
###############################################################################################################


#sub routine to return an array of randomly generated arguments.
sub genNumbers {
  my $random_num = 2 + int(rand(100));

  my @arr = ();
  my $i = 0;

  while ($i < $random_num) {
    my $num = int(rand(10000000));
    $arr[$i] = $num;
    $i++;
  }

  return @arr;
}




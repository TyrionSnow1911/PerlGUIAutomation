use strict;
use warnings;
use Win32;
use Win32::API;
use Win32::OLE;
use Cwd;
my $Au3 = Win32::OLE->new ("AutoItX3.Control"); #Slave AutoItX3 with perl.

my $rootDir = cwd;
my $testDir = "$rootDir\\Test_Cases";

#######################Generate addition test input file.###############################
open my $fileHandle, '>',
    "$testDir\\addition_test_input_file.txt" or die
        "Cannot create test case file: $testDir\\addition_test_input_file.txt!";
my $numOfTestCases = 3;

while ($numOfTestCases-- > 0) {
  my @numbers = genNumbers();
  addArrayelements($fileHandle,\@numbers);#create a string from numbers array in the form a+b+c...=d
}
close $fileHandle;
#######################################################################################


#######Generate text file of expected outputs from addition_test_input_file.txt########
open my $_fileHandle, '<',
    "$testDir\\addition_test_input_file.txt" or die
        "Cannot create test case file: $testDir\\addition_test_input_file.txt!";

open my $FH, '>', "$testDir\\expected_addition_test_output.txt" or die
        "Cannot create test case file: $testDir\\expected_addition_test_output.txt!";

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
#Extract input arguments from input test file.
#open addition test input file for reading.
open my $input, '<',
    "$testDir\\addition_test_input_file.txt" or die
        "Cannot create test case file: $testDir\\addition_test_input_file.txt!";
#open addition test output file for writing.
open my $output, '>',
    "$testDir\\actual_addition_test_output.txt" or die
        "Cannot create test case file: $testDir\\actual_addition_test_output.txt!";

$Au3->ClipGet(); #clear clip board cache.	
$Au3->ClipGet();
$Au3->ClipGet();	
$Au3->Run("calc.exe");#Open Calculator
		
				
while (my $operation = <$input>) {
chomp $operation;

$Au3->WinWaitActive("Calculator");
$Au3->WinWaitActive("Calculator");#Bring Calculator to focus twice.


my @operands = split /\+|\=/, $operation;
pop @operands;

my $res = driveCalculator(\@operands);
print $output $res;
print $output "\n";
#last;

# print join ",",@operands;
# print "\n\n";
}

close $input;
close $output;

$Au3->Sleep(500); # wait 5 secs 
$Au3->WinClose("Calculator");# close calculator


#sub-routine to run each test case through Calculator GUI and write result to output file.
sub driveCalculator {
  my $arr_ref = shift;
  my @numbers = @$arr_ref;
  my $arr_size = scalar @numbers;

  my $ind = 0;
  
  while ($ind < $arr_size) {
  
	my @digits = split //, $numbers[$ind]; #convert num/string into an array of chars.
	my $_n = scalar @digits;
	for(my $k = 0 ; $k < $_n; $k++) {#transmit one digit at a time to calculator.
		#print "$digits[$k]";
		$Au3->WinWaitActive("Calculator");
		$Au3->Sleep(400);
		$Au3->Send("{$digits[$k]}");
    }
	#system("pause");
	$ind++;

    if ($ind < $arr_size) {#if not last input, send a plus sign to GUI.
		$Au3->Send("{+}");
    }
	else {#if reached last input, then hit Enter.
		$Au3->Send("{=}");
	}
  }
  #$Au3->Sleep(400);
  
  #extract result from Calculator GUI, wait 5 secs then close.
  $Au3->Send("{CTRLDOWN}c");# use ctrl + c to capture calculator result and store in sys cache.
  $Au3->Sleep(400);
  $Au3->Send("{CTRLUP}");
  $Au3-> ClipPut();
  $Au3->Sleep(400);
  my $result  = $Au3->ClipGet();
  $Au3->Sleep(400);  
  $Au3->Send("{ESC}");
  
  return $result;
}



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
  my $random_num = 2 + int(rand(3)); #randomly determine array size.

  my @arr = ();
  my $i = 0;

  while ($i < $random_num) {
    my $num = int(rand(10000000));
    $arr[$i] = $num;
    $i++;
  }

  return @arr;
}




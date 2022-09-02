#!/usr/bin/perl
use File::Path qw( rmtree );
=pod
sub install_module(){
    print "Press ENTER to Install...";
    <STDIN>;
    if($^O == "linux"){
        system ("sudo apt-get install cpanminus -y"); # libpath-tiny-Perl
    }
    system ("cpan install File::Copy");
    system ("cpan install Archive::Zip");
    system ("cpan install Archive::Extract");
}
=cut
BEGIN {
    my $fails;
    foreach my $module ( qw/File::Copy Archive::Zip Archive::Extract/ ) {
        eval {
            eval "require $module" or die;
            $module->import;
        };
        if ($@ && $@ =~ /$module/) {
            warn "You need to install the $module module";
            $fails++;
        }
    }
    #install_module();
    exit if $fails;
}

my ($path_to_docx_file) = @ARGV;

$path_to_docx_file_new      = (substr $path_to_docx_file, 0, rindex( $path_to_docx_file, q{.} )) . "_new.docx";

$path_to_zip_file   = $path_to_docx_file . ".zip";
$path_to_unzip      = $path_to_docx_file . ".tmp";
$docx_xml_loc       = $path_to_unzip . "\\word\\document.xml";
my @docx_xml;
my $regex = qr/<w:t xml:space="preserve">Finding XX\:/mp;

copy($path_to_docx_file, $path_to_zip_file);

my $ae = Archive::Extract->new( archive => $path_to_zip_file );
my $ok = $ae->extract( to => $path_to_unzip );
my $files   = $ae->files;
my $outdir  = $ae->extract_path;
unlink($path_to_zip_file) or die "Can't unlink $file: $!";
open(my $fh, '<:encoding(UTF-8)', $docx_xml_loc) or die "Unable to open file!";
my @docx_xml = <$fh>;
$str = join("", @docx_xml);
my $totalFinding = () = $str =~ /$regex/gi;
for($i=1;$i<=$totalFinding;$i++){
    if($i<10){
        ($str =~ s/$regex/<w:t xml:space="preserve">Finding 0$i\:/i);
    }else{
        ($str =~ s/$regex/<w:t xml:space="preserve">Finding $i\:/i);
    }
    ($str =~ s/&lt;p&gt;//i);
    ($str =~ s/&lt;\/p&gt;//i);
}
close $fh;
open (FILE, "> " . $docx_xml_loc) || die "problem writing file\n"; 
binmode(FILE, ":utf8");
print FILE $str;
close FILE;

my $base_folder = ( grep -d, glob $path_to_unzip)[-1];
my $zip = Archive::Zip->new;
$zip->addTree({
  root => $base_folder
});
$zip->writeToFileNamed($path_to_docx_file_new) == AZ_OK or die "Failed to save zip file: $!";
sleep(3);
rmtree( $path_to_unzip ) or die "\nCouldn't remove $dir directory, $!";
print("\t[x] Docx file written");


=pod
# TODO
my $regex_hyperlinkand_no = qr/<w:t xml:space="preserve">(\d+)<\/w:t><\/w:r><\/w:p><\/w:tc><w:tc><w:tcPr><w:tcW w:w="(\d+)" w:type="pct"\/><\/w:tcPr><w:p w14:paraId="(\w{8})" w14:textId="(\w{8})" w:rsidR="(\w{8})" w:rsidRPr="(\w{8})" w:rsidRDefault="(\w{8})" w:rsidP="(\w{8})"><w:pPr><w:rPr><w:rFonts w:ascii="(.*?)" w:hAnsi="(.*?)"\/><w:b w:val="(\d+)"\/><w:bCs\/><w:sz w:val="(\d+)"\/><w:szCs w:val="(\d+)"\/><\/w:rPr><\/w:pPr><w:proofErr w:type="gramStart"\/><w:r w:rsidRPr="(\w{8})"><w:rPr><w:rFonts w:ascii="(.*?)" w:hAnsi="(.*?)"\/><w:b w:val="(\d+)"\/><w:sz w:val="(\d+)"\/><w:szCs w:val="(\d+)"\/><\/w:rPr><w:t xml:space="preserve">(.*?)<\/w:t>/mp;
my $replace_finding_no = "<w:t xml:space=\"preserve\">$index<\/w:t><\/w:r><\/w:p><\/w:tc><w:tc><w:tcPr><w:tcW w:w=\"$2\" w:type=\"pct\"\/><\/w:tcPr><w:p w14:paraId=\"$3\" w14:textId=\"$4\" w:rsidR=\"$5\" w:rsidRPr=\"$6\" w:rsidRDefault=\"$7\" w:rsidP=\"$8\"><w:pPr><w:rPr><w:rFonts w:ascii=\"$9\" w:hAnsi=\"$10\"\/><w:b w:val=\"$11\"\/><w:bCs\/><w:sz w:val=\"$12\"\/><w:szCs w:val=\"$13\"\/><\/w:rPr><\/w:pPr><w:proofErr w:type=\"gramStart\"\/><w:r w:rsidRPr=\"$14\"><w:rPr><w:rFonts w:ascii=\"$15\" w:hAnsi=\"$16\"\/><w:b w:val=\"$17\"\/><w:sz w:val=\"$18\"\/><w:szCs w:val=\"$19\"\/><\/w:rPr><w:t xml:space=\"preserve\">$20<\/w:t>";
=cut

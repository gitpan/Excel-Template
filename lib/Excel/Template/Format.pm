package Excel::Template::Format;

use strict;

# This is the format repository. Spreadsheet::WriteExcel does not cache the
# known formats. So, it is very possible to continually add the same format
# over and over until you run out of RAM or addressability in the XLS file. In
# real life, less than 10-20 formats are used, and they're re-used in various
# places in the file. This provides a way of keeping track of already-allocated
# formats and making new formats based on old ones.

my %_Parameters = do {
    my $i = 0;
    (map { $_ => $i++ } qw(
        bold italic
    ));
};

sub _params_to_vec
{
    my %params = @_;

    my $vec = '';

    vec($vec, $_Parameters{$_}, 1) = 1
        for grep { exists $_Parameters{$_} }
            map { lc } keys %params;

    $vec;
}

sub _vec_to_params
{
    my ($vec) = @_;

    my %params;
    while (my ($k, $v) = each %_Parameters)
    {
        next unless vec($vec, $v, 1);
        $params{$k} = 1;
    }

    %params;
}

my %_Formats;

sub _assign {
    $_Formats{$_[0]} = $_[1] unless exists $_Formats{$_[0]};
    $_Formats{$_[1]} = $_[0] unless exists $_Formats{$_[1]};
}

sub _retrieve_vec    { ref($_[0]) ? ($_Formats{$_[0]}) : ($_[0]); }
sub _retrieve_format { ref($_[0]) ? ($_[0]) : ($_Formats{$_[0]}); }

sub blank_format
{
    shift;
    my ($context) = @_;

    my $blank_vec = _params_to_vec();

    my $format = _retrieve_format($blank_vec);
    return $format if $format;

    $format = $context->{XLS}->add_format;
    _assign($blank_vec, $format);
    $format;
}

sub copy
{
    shift;
    my ($context, $old_format, %properties) = @_;

    defined(my $vec = _retrieve_vec($old_format))
        || die "Internal Error: Cannot find vector for format '$old_format'!\n";

    my $new_vec = _params_to_vec(%properties);

    $new_vec |= $vec;

    my $format = _retrieve_format($new_vec);
    return $format if $format;

    $format = $context->{XLS}->add_format(_vec_to_params($new_vec));
    _assign($new_vec, $format);
    $format;
}

1;
__END__


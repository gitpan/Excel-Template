package Excel::Template::Factory;

use strict;

BEGIN {
    use vars qw(%Manifest %isBuildable);
}

%Manifest = (

# These are the instantiable nodes
    'IF'        => 'Excel::Template::Container::Conditional',
    'LOOP'      => 'Excel::Template::Container::Loop',
    'ROW'       => 'Excel::Template::Container::Row',
    'SCOPE'     => 'Excel::Template::Container::Scope',
    'WORKBOOK'  => 'Excel::Template::Container::Workbook',
    'WORKSHEET' => 'Excel::Template::Container::Worksheet',

    'CELL'      => 'Excel::Template::Element::Cell',
    'VAR'       => 'Excel::Template::Element::Var',

    'BOLD'      => 'Excel::Template::Container::Bold',
    'ITALIC'    => 'Excel::Template::Container::Italic',
#    'FONT'      => 'Excel::Template::Container::Font',

# These are the helper objects

    'CONTEXT'    => 'Excel::Template::Context',
    'ITERATOR'   => 'Excel::Template::Iterator',
    'TEXTOBJECT' => 'Excel::Template::TextObject',

    'CONTAINER'  => 'Excel::Template::Container',
    'ELEMENT'    => 'Excel::Template::Element',

    'BASE'       => 'Excel::Template::Base',
);

while (my ($k, $v) = each %Manifest)
{
    (my $n = $v) =~ s!::!/!g;
    $n .= '.pm';

    $Manifest{$k} = {
        package  => $v,
        filename => $n,
    };
}

%isBuildable = map { $_ => 1 } qw(
    CELL
    BOLD
    IF
    ITALIC
    LOOP
    ROW
    VAR
    WORKBOOK
    WORKSHEET
);

sub register
{
    my %params = @_;

    my @param_names = qw(name class isa);
    for (@param_names)
    {
        unless ($params{$_})
        {
            warn "$_ was not supplied to register()\n";
            return 0;
        }
    }

    my $name = uc $params{name};
    if (exists $Manifest{$name})
    {
        warn "$params{name} already exists in the manifest.\n";
        return 0;
    }

    my $isa = uc $params{isa};
    unless (exists $Manifest{$isa})
    {
        warn "$params{isa} does not exist in the manifest.\n";
        return 0;
    }

    $Manifest{$name} = $params{class};
    $isBuildable{$name} = 1;

    {
        no strict 'refs';
        unshift @{"$params{class}::ISA"}, $Manifest{$isa};
    }

    return 1;
}

sub create
{
    my $class = shift;
    my $name = uc shift;

    return unless exists $Manifest{$name};

    eval {
        require $Manifest{$name}{filename};
    }; if ($@) {
        print "$@\n";
        die "Cannot find PM file for '$name' ($Manifest{$name}{filename})\n";
    }

    return $Manifest{$name}{package}->new(@_);
}

sub create_node
{
    my $class = shift;
    my $name = uc shift;

    return unless exists $isBuildable{$name};

    return $class->create($name, @_);
}

sub isa
{
    return unless @_ >= 2;
    exists $Manifest{uc $_[1]}
        ? UNIVERSAL::isa($_[0], $Manifest{uc $_[1]}{package})
        : UNIVERSAL::isa(@_)
}

1;
__END__

=head1 NAME

Excel::Template::Factory

=head1 PURPOSE

=head1 NODE NAME

=head1 INHERITANCE

=head1 ATTRIBUTES

=head1 CHILDREN

=head1 AFFECTS

=head1 DEPENDENCIES

=head1 USAGE

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

=cut

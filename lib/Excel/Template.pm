package Excel::Template;

use strict;

BEGIN {
    use Excel::Template::Base;
    use vars qw ($VERSION @ISA);

    $VERSION  = 0.01;
    @ISA      = qw (Excel::Template::Base);
}

use XML::Parser;
use Spreadsheet::WriteExcel;

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    $self->parse_xml($self->{FILENAME})
        if defined $self->{FILENAME};

    return $self;
}

sub param
{
    my $self = shift;

    # Allow an arbitrary number of hashrefs, so long as they're the first things    # into param(). Put each one onto the end, de-referenced.
    push @_, %{shift @_} while UNIVERSAL::isa($_[0], 'HASH');

    (@_ % 2)
        && die __PACKAGE__, "->param() : Odd number of parameters to param()\n";
    my %x = @_;
    @{$self->{PARAM_MAP}}{keys %x} = @x{keys %x};

    return 1;
}

sub write_file
{
    my $self = shift;
    my ($filename) = @_;

    my $xls = Spreadsheet::WriteExcel->new($filename)
        || die "Cannot create XLS in '$filename': $!\n";

    $self->_prepare_output($xls);

    $xls->close;

    return 1;
}

sub output
{
    my $self = shift;

    # mod_perl 2.x
    if ( $ENV{MOD_PERL} && 0)
    {
        eval {
            require Apache::IO;
            Apache::IO->import;
        }; if ($@) {
            die "Cannot figure out what to do under Apache!\n";
        }

        tie *XLS, 'Apache::IO';
        binmode *XLS;
        return $self->write_file(\*XLS);
    }
    # mod_perl 1.x (or CGI under mod_perl 2.x)
    elsif ( $ENV{GATEWAY_INTERFACE} && 0)
    {
        my $module_name;
        eval {
           require Apache;
           Apache->import;
           $module_name = 'Apache';
        }; if ($@) {
            die "Cannot figure out what to do under Apache!\n";
        }

        tie *XLS, $module_name;
        binmode *XLS;
        return $self->write_file(\*XLS);
    }
    # CGI (or non-WWW) behavior
    else
    {
        $self->write_file('-');
    }
}

sub parse
{
    my $self = shift;

    $self->parse_xml(@_);
}

sub parse_xml
{
    my $self = shift;
    my ($filename) = @_;

    my @stack;
    my $parser = XML::Parser->new(
        Handlers => {
            Start => sub {
                shift;

                my $name = uc shift;

                my $node = Excel::Template::Factory->create_node($name, @_);
                die "'$name' (@_) didn't make a node!\n" unless defined $node;

                if ($name eq 'WORKBOOK')
                {
                    push @{$self->{WORKBOOKS}}, $node;
                }
                elsif ($name eq 'VAR')
                {
                    return unless @stack;
                                                                                
                    if (exists $stack[-1]{TXTOBJ} &&
                        $stack[-1]{TXTOBJ}->isa('TEXTOBJECT'))
                    {
                        push @{$stack[-1]{TXTOBJ}{STACK}}, $node;
                    }
 
                }
                else
                {
                    push @{$stack[-1]{ELEMENTS}}, $node
                        if @stack;
                }
                push @stack, $node;
            },
            Char => sub {
                shift;
                return unless @stack;

                my $parent = $stack[-1];

                if (
                    exists $parent->{TXTOBJ}
                        &&
                    $parent->{TXTOBJ}->isa('TEXTOBJECT')
                ) {
                    push @{$parent->{TXTOBJ}{STACK}}, @_;
                }
            },
            End => sub {
                shift;
                return unless @stack;

                pop @stack if $stack[-1]->isa(uc $_[0]);
            },
        },
    );

    {
        my $fh = IO::File->new($filename)
            || die "Cannot open '$filename' for reading: $!\n";

        $parser->parse(do { local $/ = undef; <$fh> });

        $fh->close;
    }

    return 1;
}

sub _prepare_output
{
    my $self = shift;
    my ($xls) = @_;

    my $context = Excel::Template::Factory->create(
        'CONTEXT',

        XLS       => $xls,
        PARAM_MAP => [ $self->{PARAM_MAP} ],
    );

    $_->render($context) for @{$self->{WORKBOOKS}};

    return 1;
}

sub register { shift; Excel::Template::Factory::register(@_) }

1;
__END__

=head1 NAME

Excel::Template - Excel::Template

=head1 SYNOPSIS

  use Excel::Template

  my $template = Excel::Template->new(
      filename => 'template.xml',
  );

  $template->param(%some_params);

  $template->write_file('output.xls');


=head1 DESCRIPTION

This is a module used for templating Excel files. Its genesis came from the
need to use the same datastructure as HTML::Template, but provide Excel files
instead. The existing modules don't do the trick, as they require separate
logic from what HTML::Template needs.

Currently, only a small subset of the planned features are supported. This is
meant to be a test of the waters, to see what features people actually want.

=head1 USAGE

=head2 new()

This creates a Excel::Template object. If passed a filename parameter, it will
parse the template in the given file. (You can also use the parse() method,
described below.)

=head2 param()

This method is exactly like HTML::Template's param() method. Although, I will
be adding more to this section later, please see HTML::Template's description
for info right now.

=head2 parse() / parse_xml()

This method actually parses the template file. It can either be called
separately or through the new() call. It will die() if it cannot handle any
situation.

=head2 write_file()

Create the Excel file and write it to the specified filename. This is when the
actual merging of the template and the parameters occurs.

=head2 output()

It will act just like HTML::Template's output() method, returning the resultant
file as a stream, usually for output to the web.

NOTE: This method binmode()'s STDOUT and passes a reference to the glob to
write_file(). This has been tested using CGI, but not mod_perl 1 or mod_perl 2.
If you test it under those environments, please let me know.

=head1 SUPPORTED NODES

This is just a list of nodes. See the other classes in this distro for more
details on specific parameters and the like.

Every node can set the ROW and COL parameters. These are the actual ROW/COL
values that the next CELL tag will write into.

=over 4

=item * WORKBOOK

=item * WORKSHEET

=item * LOOP

=item * ROW

=item * CELL

=item * BOLD

=item * IF

=back 4

=head1 BUGS

None, that I know of. (But there aren't many features, neither!)

=head1 SUPPORT

This is currently beta-quality software. The featureset is extremely limited,
but I expect to be adding on to it very soon.

If you have any suggestions, comments, or the like, please let me know. I'd
also appeciate anyone coming up with a good testing strategy!

=head1 AUTHOR

    Rob Kinyon
    rkinyon@columbus.rr.com

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

perl(1), HTML::Template, PDF::Template.

=cut

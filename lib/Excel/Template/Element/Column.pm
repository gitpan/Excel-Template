package Excel::Template::Element::Column;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Element);

    use Excel::Template::Element;
}

sub render
{
    my $self = shift;
    my ($context) = @_;

    $context->active_worksheet->set_column(
        (map { $context->get($self, $_) } qw( NAME WIDTH )),
    );

    return 1;
}

1;
__END__

=head1 NAME

Excel::Template::Element::Column - Excel::Template::Element::Column

=head1 PURPOSE

To set the width of a column or columns

=head1 NODE NAME

COLUMN

=head1 INHERITANCE

Excel::Template::Element

=head1 ATTRIBUTES

=over 4

=item * NAME

This is the Excel name for the column(s) you want to affect. This is a range, so
you will need to say "A:A" in order to affect the first column.

=item * WIDTH

This is the width you want to set the column to. To determine a size, you should
figure on roughly how long you expect a string in default Arial 10 to be. Since
Auto-fit is a feature only available from within Excel at runtime, there is no
way to specify it through either this distribution or [Spreadheet::WriteExcel]
(the rendering engine).

=back 4

There will be more parameters added, as features are added. Additionally, this
node might (someday) be able to calculate how much space the data you have
written to that column actually needs.

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

  <column name="A:A" width="20" />
  <column name="C:E" width="30" />

In the above example, the width of the first column is set to 20 and the width
of the third, fourth, and fifth columns are set to 30.

This node can be used anywhere under a worksheet node.

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

Nothing

=cut

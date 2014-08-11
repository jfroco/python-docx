
Table - Merge Cells
===================

Word allows contiguous table cells to be merged, such that two or more cells
appear to be a single cell. Cells can be merged horizontally (spanning
multple columns) or vertically (spanning multiple rows). Cells can also be
merged both horizontally and vertically at the same time, producing a cell
that spans both rows and columns. Only rectangular ranges of cells can be
merged.


Cell collection behavior
------------------------

Cell collection length is unaffected by the presence of merged cells.
However, indexed access to a cell behaves a little differently when merged
cells are involved. The range of valid row and column indicies is determined
by the layout grid and is unaffected by spans. This makes *continuation*
cells addressable along with uniform and *origin* cells. All layout grid
(row, col) addresses return a cell reference, however a continuation cell
address returns its corresponding origin cell. An origin cell will have as
many addresses as layout cells that compose it.


Cell access protocol
--------------------

Two-cell horizontal merge in uniform 3 x 3 table::

    >>> table = document.add_table(3, 3)
    >>> row = table.rows[0]
    >>> len(row.cells)
    3
    >>> cell, cell_2 = row.cells[:2]
    >>> cell.width.inches, cell_2.width.inches
    (1.0, 1.0)
    >>> cell.merge(cell_2)
    >>> len(row.cells)
    3
    >>> cell.width.inches
    2.0
    >>> row.cells[0] == row.cells[1]
    True
    >>> cell_2.width
    SomeWierdError ... invalid element reference


Acceptance tests
----------------

* Add feature file for tbl-cell-access.feature, moving relevant parts from
  tbl-item-access.feature. Rename item to tbl-coll-access.feature.

* Add scenarios for len and indexed access for horizontal, vertical, and
  combined spans.

  + len x uniform
  + len x horz span
  + len x vert span
  + len x comb span


::

  Given a 3x3 table containing <span-state>
   Then len(table.columns) is <cols>
    And len(table.rows) is <rows>
    And len(row.cells) is <cols> for each of its rows
    And len(column.cells) is <rows> for each of its columns

  Examples: ...
    | span-state         | rows | cols |
    | only uniform cells |   3  |   3  |
    | a horizontal span  |   3  |   3  |
    | a vertical span    |   3  |   3  |
    | a combined span    |   3  |   3  |


  Given a row with a horizontal span of its first two cells
   Then row.cells[0] == row.cells[1]


Cell access notes
-----------------

There are three ways to access a table cell in the MS API:

* Table.cell(row_idx, col_idx)
* Row.cells[col_idx]
* Column.cells[col_idx]

Table.cell() is really just shorthand for Table.rows[row_idx].cells[col_idx].


Indexed access
~~~~~~~~~~~~~~

`len()`
~~~~~~~

`len()` always bases its count on the layout grid, as though there were no
merged cells.

* ``len(Table.columns)`` is the number of `w:gridCol` elements, representing
  the number of grid columns, without regard to the presence of merged cells
  in the table.

* ``len(Table.rows)`` is the number of `w:tr` elements, regardless of any
  merged cells that may be present in the table.

* ``len(Row.cells)`` is the number of grid columns, regardless of whether any
  cells in the row are merged.

* ``len(Column.cells)`` is the number of rows in the table, regardless of
  whether any cells in the column are merged.


Protocol notes
--------------

* Cell.__eq__ -> self._element is other._element
* Any extra cell iterators required? Not seeing the use case yet. Although
  implementing as an iterator and using list(iter) to provide sequences is
  a flexible approach.


Glossary
--------

layout grid
    The regular two-dimensional matrix of rows and columns that determines
    the layout of cells in the table. The grid is primarily defined by the
    `w:gridCol` elements that define the layout columns for the table. Each
    row essentially duplicates that layout for an additional row, although
    its height can differ from other rows. Every actual cell in the table
    must begin and end on a layout grid "line", whether the cell is merged or
    not.

span
    The single "combined" cell occupying the area of a set of merged cells.

skipped cell
    The WordprocessingML (WML) spec allows for 'skipped' cells, where
    a layout cell location contains no actual cell. I can't find a way to
    make a table like this using the Word UI and haven't experimented yet to
    see whether Word will load one constructed by hand in the XML.

uniform table
    A table in which each cell corresponds exactly to a layout cell.
    A uniform table contains no spans or skipped cells.

non-uniform table
    A table that contains one or more spans, such that not every cell
    corresponds to a single layout cell. I suppose it would apply when there
    was one or more skipped cells too, but in this analysis the term is only
    used to indicate a table with one or more spans.

uniform cell
    A cell not part of a span, occupying a single cell in the layout grid.

origin cell
    The top-leftmost cell in a span. Contrast with *continuation cell*.

continuation cell
    A layout cell that has been subsumed into a span. A continuation cell is
    mostly an abstract concept, although a actual `w:tc` element will always
    exist in the XML for each continuation cell in a vertical span.


Open Issues
-----------

Does it account for "skipped" cells at the beginning of a row (`w:gridBefore`
element)?


Word behavior
-------------

* Row and Column access in the MS API just plain breaks when the table is not
  uniform. `Table.Rows(n)` and `Cell.Row` raise an `EnvironmentError` when a
  table contains a vertical span, and `Table.Columns(n)` and `Cell.Column`
  unconditionally raise an EnvironmentError when the table contains
  a horizontal span.

  I'm pretty sure we can do better.

* `Table.Cell(n, m)` works on any non-uniform table, although it uses
  a *visual grid* that greatly complicates access. It raises an error for `n`
  or `m` out of visual range, and provides no way other than try/except to
  determine what that visual range is, since `Row.Count` and `Column.Count`
  are unavailable.

* In a simple 2-cell vertical or horizontal merge operation, the text of the
  continuation cell is appended to that of the origin cell as a separate
  paragraph(s).

* If a merge range contains previously merged cells, the range must
  completely enclose the merged cells.

* Vertically merged cells marked by ``w:vMerge=continue`` are not accessible
  via the MS API. Attempting to access a "continuation" cell raises an
  exception with the message "The member of the collection does not exist".

* Horizontally merged cells other than the leftmost are deleted and cannot
  be accessed via the MS API.

* Word resizes a table (adds rows) when a cell is referenced by an
  out-of-bounds row index. If the column identifier is out of bounds, an
  exception is raised.

* An exception is raised when attempting to merge cells from different tables.


XML Semantics
-------------

In a horizontal merge, the ``<w:tc w:gridSpan="?">`` attribute indicates the
number of columns the cell should span. Only the leftmost cell is preserved;
the remaining cells in the merge are deleted.

For merging vertically, the ``w:vMerge`` table cell property of the uppermost
cell of the column is set to the value "restart" of type ``w:ST_Merge``. The
following, lower cells included in the vertical merge must have the
``w:vMerge`` element present in their cell property (``w:TcPr``) element. Its
value should be set to "continue", although it is not necessary to
explicitely define it, as it is the default value. A vertical merge ends as
soon as a cell ``w:TcPr`` element lacks the ``w:vMerge`` element. Similarly
to the ``w:gridSpan`` element, the ``w:vMerge`` elements are only required
when the table's layout is not uniform across its different columns. In the
case it is, only the topmost cell is kept; the other lower cells in the
merged area are deleted along with their ``w:vMerge`` elements and the
``w:trHeight`` table row property is used to specify the combined height of
the merged cells.


Algorithm
---------

**Collapsing a column.** When all rows in a table share the same
``w:gridSpan`` specification, the spanned columns can be collapsed into
a single column of their combined width.


.. python-docx API refinements over Word's
.. ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. Addressing some of the Word API deficiencies when dealing with merged cells,
.. the following new features were introduced:

.. * A row or column has a defined length when it contains merged cells. 
..   The reported length includes the normal (unmerged) cells, plus all the
..   *master* merged cells. By *master* merged cells, we understand the leftmost
..   cell of an horizontally merged area, the top-most cell of a vertically
..   merged area, or the top-left-most cell of two-ways merged area.

.. * The same logic is applied to filter the iterable cells in a _ColumnCells or
..   _RowCells cells collection and a restricted access error message is written
..   when trying to access visually hidden, non master merged cells.

.. .. note:: Not liking this next idea yet. Or basically this notion of
..    "restricted access cells" in general. What is the purpose of accessing
..    a merge "continuation" cell? All the action is happening in the "master"
..    cell, isn't it? Plus what do you do if no continuation cell exists because
..    a `w:gridSpan` element was used without a `w:hMerge` element? That would
..    make it a "maybe there" cell, which would complicate client code like
..    crazy.

.. * The smart filtering of hidden merged cells, dubbed *visual grid* can be
..   turned off to gain access to cells which would normally be restricted,
..   either via the ``Table.cell`` method's third argument, or by setting the
..   ``visual_grid`` static property of a ``_RowCells`` or ``_ColumnsCell``
..   instance to *False*.


Candidate protocol -- cell.merge()
----------------------------------

The following interactive session demonstrates the protocol for merging table
cells. The capability of reporting the length of merged cells collection is
also demonstrated::

    >>> table = document.add_table(5, 5)
    >>> table.cell(0, 0).merge(table.cell(3, 3))
    >>> len(table.columns[2].cells)
    1
    >>> cells = table.columns[2].cells
    >>> cells.visual_grid = False
    >>> len(cells)
    5

Specimen XML
------------

.. highlight:: xml

A 3 x 3 table where an area defined by the 2 x 2 topleft cells has been
merged, demonstrating the combined use of the ``w:gridSpan`` as well as the
``w:vMerge`` elements, as produced by Word::

  <w:tbl>
    <w:tblPr>
       <w:tblW w:w="0" w:type="auto" />
    </w:tblPr>
    <w:tblGrid>
       <w:gridCol w:w="3192" />
       <w:gridCol w:w="3192" />
       <w:gridCol w:w="3192" />
    </w:tblGrid>
    <w:tr>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="6384" w:type="dxa" />
             <w:gridSpan w:val="2" />
             <w:vMerge w:val="restart" />
          </w:tcPr>
       </w:tc>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="3192" w:type="dxa" />
          </w:tcPr>
       </w:tc>
    </w:tr>
    <w:tr>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="6384" w:type="dxa" />
             <w:gridSpan w:val="2" />
             <w:vMerge />
          </w:tcPr>
       </w:tc>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="3192" w:type="dxa" />
          </w:tcPr>
       </w:tc>
    </w:tr>
    <w:tr>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="3192" w:type="dxa" />
          </w:tcPr>
       </w:tc>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="3192" w:type="dxa" />
          </w:tcPr>
       </w:tc>
       <w:tc>
          <w:tcPr>
             <w:tcW w:w="3192" w:type="dxa" />
          </w:tcPr>
       </w:tc>
    </w:tr>
  </w:tbl>


Schema excerpt
--------------

.. highlight:: xml

::

  <xsd:complexType name="CT_TcPr">  <!-- denormalized -->
    <xsd:sequence>
      <xsd:element name="cnfStyle"             type="CT_Cnf"           minOccurs="0"/>
      <xsd:element name="tcW"                  type="CT_TblWidth"      minOccurs="0"/>
      <xsd:element name="gridSpan"             type="CT_DecimalNumber" minOccurs="0"/>
      <xsd:element name="hMerge"               type="CT_HMerge"        minOccurs="0"/>
      <xsd:element name="vMerge"               type="CT_VMerge"        minOccurs="0"/>
      <xsd:element name="tcBorders"            type="CT_TcBorders"     minOccurs="0"/>
      <xsd:element name="shd"                  type="CT_Shd"           minOccurs="0"/>
      <xsd:element name="noWrap"               type="CT_OnOff"         minOccurs="0"/>
      <xsd:element name="tcMar"                type="CT_TcMar"         minOccurs="0"/>
      <xsd:element name="textDirection"        type="CT_TextDirection" minOccurs="0"/>
      <xsd:element name="tcFitText"            type="CT_OnOff"         minOccurs="0"/>
      <xsd:element name="vAlign"               type="CT_VerticalJc"    minOccurs="0"/>
      <xsd:element name="hideMark"             type="CT_OnOff"         minOccurs="0"/>
      <xsd:element name="headers"              type="CT_Headers"       minOccurs="0"/>
      <xsd:choice                                                      minOccurs="0"/>
        <xsd:element name="cellIns"            type="CT_TrackChange"/>
        <xsd:element name="cellDel"            type="CT_TrackChange"/>
        <xsd:element name="cellMerge"          type="CT_CellMergeTrackChange"/>
      </xsd:choice>
      <xsd:element name="tcPrChange"           type="CT_TcPrChange"    minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>

  <xsd:complexType name="CT_DecimalNumber">
    <xsd:attribute name="val" type="ST_DecimalNumber" use="required"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_DecimalNumber">
     <xsd:restriction base="xsd:integer"/>
  </xsd:simpleType>

  <xsd:complexType name="CT_VMerge">
    <xsd:attribute name="val" type="ST_Merge"/>
  </xsd:complexType>

  <xsd:complexType name="CT_HMerge">
    <xsd:attribute name="val" type="ST_Merge"/>
  </xsd:complexType>

  <xsd:simpleType name="ST_Merge">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="continue"/>
      <xsd:enumeration value="restart"/>
    </xsd:restriction>
  </xsd:simpleType>


Ressources
----------

* `Cell.Merge Method on MSDN`_

.. _`Cell.Merge Method on MSDN`:
   http://msdn.microsoft.com/en-us/library/office/ff821310%28v=office.15%29.aspx

Relevant sections in the ISO Spec
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
* 17.4.17 gridSpan (Grid Columns Spanned by Current Table Cell)
* 17.4.84 vMerge (Vertically Merged Cell)
* 17.18.57 ST_Merge (Merged Cell Type)

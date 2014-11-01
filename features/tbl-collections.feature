Feature: Access table objects using its collections
  In order to access individual rows, columns, and cells in a table
  As an python-docx developer
  I need a set of table object collections


  @wip
  Scenario Outline: Get collection length
    Given a 3x3 table containing <span-state>
     Then len(table.columns) is <cols>
      And len(table.rows) is <rows>
      And len(row.cells) is <cols> for each of its rows
      And len(column.cells) is <rows> for each of its columns

    Examples: Tables having varied spans
      | span-state         | rows | cols |
      | only uniform cells |   3  |   3  |
      | a horizontal span  |   3  |   3  |
      | a vertical span    |   3  |   3  |
      | a combined span    |   3  |   3  |

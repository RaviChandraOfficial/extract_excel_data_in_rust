use calamine::{open_workbook, Data, Reader, Xlsx};

fn main() {
    // Specify the path to your Excel file
    let path = "/home/ravi/ECL2 Projects/project yashwanth/files/checking.xlsx";

    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open Excel file");

    // Specify the sheet name
    let sheet_name = "Sheet1"; // Replace with your sheet name

    // Get the range of the specified sheet and extract data
    match workbook.worksheet_range(sheet_name) {
        Ok(range) => {
            println!("Data in '{}':", sheet_name);
            for row in range.rows() {
                for cell in row {
                    print!("{:?}\t", cell);
                }
                println!();
            }

            // Provide some statistics
            let total_cells = range.get_size().0 * range.get_size().1;
            let non_empty_cells: usize = range.used_cells().count();
            println!(
                "Found {} cells in '{}', including {} non-empty cells",
                total_cells, sheet_name, non_empty_cells
            );
            // alternatively, we can manually filter rows
            assert_eq!(
                non_empty_cells,
                range.rows()
                    .flat_map(|r| r.iter().filter(|&c| c != &Data::Empty))
                    .count()
            );
        }
        Err(_) => println!("Sheet '{}' does not exist in the workbook", sheet_name),
        Err(e) => println!("Error reading the workbook: {:?}", e),
    }

    // Check if the workbook has a VBA project
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        let module1 = vba.get_module("Module 1").unwrap();
        println!("Module 1 code:");
        println!("{}", module1);
        for r in vba.get_references() {
            if r.is_missing() {
                println!("Reference {} is broken or not accessible", r.name);
            }
        }
    }

    // Get defined names definition (string representation only)
    for name in workbook.defined_names() {
        println!("Name: {}, Formula: {}", name.0, name.1);
    }

    // Get all formulas in the sheets
    let sheets = workbook.sheet_names().to_owned();
    for s in sheets {
        println!(
            "Found {} formulas in '{}'",
            workbook
                .worksheet_formula(&s)
                .expect("Error while getting formula")
                .rows()
                .flat_map(|r| r.iter().filter(|f| !f.is_empty()))
                .count(),
            s
        );
    }
}

use std::fs::File;
use std::io::{BufRead, BufReader, Write};
use calamine::{DataType, Reader, open_workbook, Xlsx};
use std::io;
use std::path::PathBuf;
use rfd::{FileDialog};
use csv::ReaderBuilder;

fn merge_data(file1: &str, file2: &str, file3: &str, file4: &str) {
    let header_text = "X(Scan ID.),Y(Channel No.),Z(Sample ID.),Amplitude(Real),Amplitude(Imaginary),Longitude,Latitude,Depth(m)";

    let mut rdr1 = ReaderBuilder::new().from_path(file1).expect("Failed to open file1");
    let mut rdr2 = ReaderBuilder::new().from_path(file2).expect("Failed to open file2");
    let mut rdr3 = ReaderBuilder::new().from_path(file3).expect("Failed to open file3");

    let mut lat_array = vec![vec![String::new(); 0]; 0];
    let mut lon_array = vec![vec![String::new(); 0]; 0];

    for result in rdr1.records() {
        let record = result.expect("Failed to read record from file1");
        let x: usize = record[0].parse().expect("Failed to parse X");
        let y: usize = record[1].parse().expect("Failed to parse Y");
        if lat_array.len() <= x {
            lat_array.resize(x + 1, vec![]);
        }
        if lat_array[x].len() <= y {
            lat_array[x].resize(y + 1, String::new());
        }
        lat_array[x][y] = record[2].to_string();
        lon_array[x][y] = record[3].to_string();
    }

    let mut df2_depth = vec![];

    for result in rdr2.records() {
        let record = result.expect("Failed to read record from file2");
        df2_depth.push(record[1].to_string());
    }

    let f = File::open(file3).expect("Failed to open file3");
    let mut out = File::create(file4).expect("Failed to create file4");
    writeln!(out, "{}", header_text).expect("Failed to write to output file");

    let reader = BufReader::new(f);
    for line in reader.lines() {
        if let Ok(line) = line {
            let parts: Vec<&str> = line.split(',').collect();
            let x: usize = parts[0].parse().expect("Failed to parse X");
            let y: usize = parts[1].parse().expect("Failed to parse Y");
            let z: usize = parts[2].parse().expect("Failed to parse Z");
            writeln!(
                out,
                "{},{},{},{},{},{},{},{}",
                parts[0],
                parts[1],
                parts[2],
                parts[3],
                parts[4],
                lon_array[x][y],
                lat_array[x][y],
                df2_depth[z]
            ).expect("Failed to write to output file");
        }
    }
}

fn main() -> io::Result<()> {

    let input_file_path = match FileDialog::new().add_filter("Excel file",&["xlsx"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };

    let input_file_path_csv = match FileDialog::new().add_filter("CVS file",&["csv"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };

    let input_file_path_latloncsv = match FileDialog::new().add_filter("Latlon CSV file",&["csv"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };

    let input_file_path_depthcsv = match FileDialog::new().add_filter("Depth CSV file",&["csv"]).pick_file() {
        Some(path) => path,
        None => return Err(io::Error::new(io::ErrorKind::Other, "No file selected.")),
    };

    //let mut str_path_csv = PathBuf::from(input_file_path_csv.parent().unwrap());


    let mut df: Vec<(String, i32, i32)> = vec![];

    // Load the Excel file
    let mut workbook: Xlsx<_> = open_workbook(input_file_path.clone()).expect("Failed to open Excel file");

    // Select the first sheet and read the data
    if let Some(Ok(r)) = workbook.worksheet_range_at(0) {
        for row in r.rows().skip(1) {
            let name = row[0].get_string().unwrap().to_string();
            let start_x = row[2].get_float().unwrap() as i32;
            let end_x = row[3].get_float().unwrap() as i32;
            df.push((name, start_x, end_x));
        }
    }

    for (i, (name, _, _)) in df.iter().enumerate() {
        println!("{}", i);
    
        let mut file1 = PathBuf::from(input_file_path_latloncsv.parent().unwrap());
        file1.push(format!(r"{} - Region1_lat_lon.csv", name));

        let file2 = input_file_path_depthcsv.to_str().unwrap();

        let mut file3 = PathBuf::from(input_file_path_csv.parent().unwrap());
        file3.push(format!(r"{}.csv", name));

        let file4 = format!(r"{}_out.csv", name);

        merge_data(&file1.to_str().unwrap(), &file2, &file3.to_str().unwrap(), &file4);
    }

    Ok(())

}


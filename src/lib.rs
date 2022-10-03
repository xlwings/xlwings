use calamine::CellErrorType::{Div0, GettingData, Name, Null, Num, Ref, Value, NA};
use calamine::{open_workbook_auto, CellErrorType, DataType, Error, Range, Reader};
use chrono::{Datelike, NaiveDateTime, Timelike};
use pyo3::exceptions::PyIOError;
use pyo3::import_exception;
use pyo3::prelude::*;
use pyo3::types::PyDateTime;

import_exception!(xlwings, XlwingsError);

struct CalamineError(Error);

impl From<CalamineError> for PyErr {
    fn from(err: CalamineError) -> PyErr {
        XlwingsError::new_err(err.0.to_string())
    }
}

impl From<Error> for CalamineError {
    fn from(other: Error) -> Self {
        Self(other)
    }
}

#[derive(Debug)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    String(String),
    DateTime(NaiveDateTime),
    Bool(bool),
    Error(CellErrorType),
    Empty,
}

impl IntoPy<PyObject> for CellValue {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            CellValue::Int(v) => v.to_object(py),
            CellValue::Float(v) => v.to_object(py),
            CellValue::String(v) => v.to_object(py),
            CellValue::Bool(v) => v.to_object(py),
            CellValue::DateTime(v) => PyDateTime::new(
                py,
                v.year(),
                v.month() as u8,
                v.day() as u8,
                v.hour() as u8,
                v.minute() as u8,
                v.second() as u8,
                v.timestamp_subsec_micros(),
                None,
            )
            .unwrap()
            .to_object(py),
            CellValue::Empty => py.None(),
            // Errors are already converted to String or Empty
            CellValue::Error(_) => String::from("Error").to_object(py),
        }
    }
}

fn get_values(
    used_range: Range<DataType>,
    cell1: (u32, u32),
    cell2: (u32, u32),
    err_to_str: bool,
) -> Result<Vec<Vec<CellValue>>, Error> {
    let mut result: Vec<Vec<CellValue>> = Vec::new();
    // println!("{:?}", range.start());
    for row in used_range.range(cell1, cell2).rows() {
        let mut result_row: Vec<CellValue> = Vec::new();

        // println!("{:?}", row);
        for value in row.iter() {
            match value {
                DataType::Int(v) => result_row.push(CellValue::Int(*v)),
                DataType::Float(v) => result_row.push(CellValue::Float(*v)),
                DataType::String(v) => result_row.push(CellValue::String(String::from(v))),
                DataType::DateTime(_v) => {
                    result_row.push(CellValue::DateTime(value.as_datetime().unwrap()))
                }
                DataType::Bool(v) => result_row.push(CellValue::Bool(*v)),
                DataType::Error(v) => match v {
                    Div0 => result_row.push(if err_to_str {
                        CellValue::String(String::from("#DIV/0!"))
                    } else {
                        CellValue::Empty
                    }),
                    NA => result_row.push(if err_to_str {
                        CellValue::String(String::from("#N/A"))
                    } else {
                        CellValue::Empty
                    }),
                    Name => result_row.push(if err_to_str {
                        CellValue::String(String::from("#NAME?"))
                    } else {
                        CellValue::Empty
                    }),
                    Null => result_row.push(if err_to_str {
                        CellValue::String(String::from("#NULL!"))
                    } else {
                        CellValue::Empty
                    }),
                    Num => result_row.push(if err_to_str {
                        CellValue::String(String::from("#NUM!"))
                    } else {
                        CellValue::Empty
                    }),
                    Ref => result_row.push(if err_to_str {
                        CellValue::String(String::from("#REF!"))
                    } else {
                        CellValue::Empty
                    }),
                    Value => result_row.push(if err_to_str {
                        CellValue::String(String::from("#VALUE!"))
                    } else {
                        CellValue::Empty
                    }),
                    GettingData => result_row.push(if err_to_str {
                        CellValue::String(String::from("#DATA!"))
                    } else {
                        CellValue::Empty
                    }),
                },
                DataType::Empty => result_row.push(CellValue::Empty),
            };
        }
        result.push(result_row);
    }
    Ok(result)
    // println!("{:?}", result)
}

#[pyfunction]
#[pyo3(text_signature = "path: str, sheet_index: int, err_to_str: bool")]
fn get_sheet_values(
    path: &str,
    sheet_index: usize,
    err_to_str: bool,
) -> PyResult<Vec<Vec<CellValue>>> {
    // TODO: proper error handling
    let mut book = open_workbook_auto(path).unwrap();
    let used_range = book.worksheet_range_at(sheet_index).unwrap().unwrap();
    let cell1 = (0, 0);
    let cell2 = match used_range.end() {
        Some((r, c)) => (r, c),
        None => (0, 0),
    };
    if used_range.is_empty() {
        return Ok(vec![vec![]]);
    }
    match get_values(used_range, cell1, cell2, err_to_str) {
        Ok(r) => Ok(r),
        Err(e) => match e {
            Error::Io(err) => Err(PyIOError::new_err(err.to_string())),
            _ => Err(XlwingsError::new_err(e.to_string())),
        },
    }
}

#[pyfunction]
#[pyo3(
    text_signature = "path: str, sheet_index: int, cell1: tuple[int, int], \
                      cell2: tuple[int, int], err_to_str: bool"
)]
fn get_range_values(
    path: &str,
    sheet_index: usize,
    cell1: (u32, u32),
    cell2: (u32, u32),
    err_to_str: bool,
) -> PyResult<Vec<Vec<CellValue>>> {
    // TODO: proper error handling
    let mut book = open_workbook_auto(path).unwrap();
    let used_range = book.worksheet_range_at(sheet_index).unwrap().unwrap();
    let used_range = match used_range.is_empty() {
        true => Range::new((0, 0), (0, 0)),
        false => used_range,
    };
    match get_values(used_range, cell1, cell2, err_to_str) {
        Ok(r) => Ok(r),
        Err(e) => match e {
            Error::Io(err) => Err(PyIOError::new_err(err.to_string())),
            _ => Err(XlwingsError::new_err(e.to_string())),
        },
    }
}

#[pyfunction]
#[pyo3(text_signature = "path: str")]
fn get_sheet_names(path: &str) -> Result<Vec<String>, CalamineError> {
    let book = open_workbook_auto(path)?;
    Ok(book.sheet_names().to_owned())
}

#[pyfunction]
#[pyo3(text_signature = "path: str")]
fn get_defined_names(path: &str) -> Result<Vec<(String, String)>, CalamineError> {
    let book = open_workbook_auto(path)?;
    Ok(book.defined_names().to_owned())
}

#[pymodule]
fn xlwingslib(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(get_range_values, m)?)?;
    m.add_function(wrap_pyfunction!(get_sheet_values, m)?)?;
    m.add_function(wrap_pyfunction!(get_sheet_names, m)?)?;
    m.add_function(wrap_pyfunction!(get_defined_names, m)?)?;
    Ok(())
}

// Based on https://github.com/dimastbk/python-calamine,
// which is released under the following license:
//
// MIT License
//
// Copyright (c) 2021 dimastbk
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

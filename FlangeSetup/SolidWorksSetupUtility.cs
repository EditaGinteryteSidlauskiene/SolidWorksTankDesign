using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;

namespace FlangeSetup
{
    /// <summary>
    /// One-time setup utility for creating SolidWorks document configurations
    /// and modifying dimensions programmatically.
    ///
    /// HOW TO USE:
    ///   1. Open the target document in SolidWorks so it becomes the active doc.
    ///   2. Call Run(doc) passing the active ModelDoc2 instance.
    ///   3. When done, remove the button and this call — or keep the file for reference.
    /// </summary>
    public static class SolidWorksSetupUtility
    {
        public static void Run(ModelDoc2 doc)
        {
            if (doc == null)
            {
                MessageBox.Show("No active document found.", "Setup Utility");
                return;
            }

            //FlangeType01Setup.ChangeDimensions(doc);

            FlangeType01Setup.ChangeDimensionsForFaces(doc);


            //doc.ForceRebuild3(true);
            //doc.Save3(
            //    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
            //    (int)swFileSaveError_e.swGenericSaveError,
            //    (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }
    }
}

/*
 * //! # Synoptic Table (Table 7)
//!
//! **Identity**:
//! This module implements the "Synoptic Table" (Table 7) from EN 1092-1.
//!
//! **Functionality**:
//! It acts as a normalization layer, mapping a requested specific flange type, Pressure Nominal (PN),
//! and Diameter Nominal (DN) to a "Target PN". This "Target PN" is then used to retrieve the
//! actual physical dimensions from the dimension tables (e.g., Table 11, Table 12).
//!
//! **Composition**:
//! *   `get_target_pn`: The main public entry point for normalization.
//! *   Internal helper functions: Optimized look tables for each flange type.
//!
//! **Usage**:
//! This module is typically called before querying specific dimension tables to ensure the
//! correct column is accessed.
//!
//! ```rust
//! use en1092_1::enums::{Dn, FlangeType, Pn};
//! use en1092_1::tables::synoptic::get_target_pn;
//!
//! let target = get_target_pn(FlangeType::Type01, Pn::Pn10, Dn::Dn40);
//! assert_eq!(target, Some(Pn::Pn40));
//! ```

use crate::enums::{Dn, FlangeType, Pn};

/// Determines the Target PN for a given flange characteristic combination.
///
/// # Arguments
///
/// * `flange_type` - The type of flange (e.g., Type 01, Type 11).
/// * `pn` - The nominal pressure (PN) designation.
/// * `dn` - The nominal diameter (DN) designation.
///
/// # Returns
///
/// Returns `Some(Pn)` containing the target PN to look up in dimension tables if the combination is valid.
/// Returns `None` if the combination is not defined in the standards or is explicitly excluded.
///
/// # Remarks
///
/// This function implements the normalization rules specified in EN 1092-1:2018, Table 7.
/// For example, a Type 01 flange with PN 10 and DN 40 uses the dimensions of PN 40.
pub const fn get_target_pn(flange_type: FlangeType, pn: Pn, dn: Dn) -> Option<Pn> {
    let dn_val = dn.value();

    match flange_type {
        FlangeType::Type01 => get_pn_type_01(pn, dn_val),
        FlangeType::Type02 => get_pn_type_02(pn, dn_val),
        FlangeType::Type05 => get_pn_type_05(pn, dn, dn_val),
        FlangeType::Type11 => get_pn_type_11(pn, dn, dn_val),
        FlangeType::Type12 => get_pn_type_12(pn, dn_val),
        FlangeType::Type13 => get_pn_type_13(pn, dn_val),
        FlangeType::Type32 => get_pn_type_32(pn, dn_val),
        FlangeType::Type33 | FlangeType::Type37 => get_pn_type_37(pn, dn_val),
        FlangeType::Type35 => get_pn_type_35(pn, dn_val),
        FlangeType::Type36 => get_pn_type_36(pn, dn_val),

        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_01(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 1000 {
                Some(Pn::Pn6)
            } else if dn_val <= 1200 {
                Some(Pn::Pn2_5)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 2000 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 1200 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 800 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 400 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        Pn::Pn63 => {
            if dn_val <= 40 {
                Some(Pn::Pn100)
            } else if dn_val <= 400 {
                Some(Pn::Pn63)
            } else {
                None
            }
        }
        Pn::Pn100 => {
            if dn_val <= 350 {
                Some(Pn::Pn100)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_32(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 600 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 600 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 600 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 600 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 600 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_02(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 1000 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 1200 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 1200 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 800 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_05(pn: Pn, dn: Dn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 1000 {
                Some(Pn::Pn6)
            } else if dn_val <= 1200 {
                Some(Pn::Pn2_5)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 2000 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 1200 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 600 {
                Some(Pn::Pn25)  // own dims
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        Pn::Pn63 => {
            if dn_val <= 40 {
                Some(Pn::Pn100)
            } else if dn_val <= 400 {
                Some(Pn::Pn63)
            } else {
                None
            }
        }
        Pn::Pn100 => {
            if dn_val <= 350 {
                Some(Pn::Pn100)
            } else {
                None
            }
        }
        Pn::Pn160 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 300 {
                    Some(Pn::Pn160)
                } else {
                    None
                }
            }
        },
        Pn::Pn250 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 10 {
                    Some(Pn::Pn320)
                } else if dn_val <= 250 {
                    Some(Pn::Pn250)
                } else {
                    None
                }
            }
        },
        Pn::Pn320 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 250 {
                    Some(Pn::Pn320)
                } else {
                    None
                }
            }
        },
        Pn::Pn400 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 200 {
                    Some(Pn::Pn400)
                } else {
                    None
                }
            }
        },
    }
}

#[inline(always)]
const fn get_pn_type_11(pn: Pn, dn: Dn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 800 {
                Some(Pn::Pn6)
            } else if dn_val <= 4000 {
                Some(Pn::Pn2_5)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 3600 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 3000 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 2000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        Pn::Pn63 => {
            if dn_val <= 40 {
                Some(Pn::Pn100)
            } else if dn_val <= 400 {
                Some(Pn::Pn63)
            } else {
                None
            }
        }
        Pn::Pn100 => {
            if dn_val <= 350 {
                Some(Pn::Pn100)
            } else {
                None
            }
        }
        Pn::Pn160 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 300 {
                    Some(Pn::Pn160)
                } else {
                    None
                }
            }
        },
        Pn::Pn250 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 10 {
                    Some(Pn::Pn320)
                } else if dn_val <= 250 {
                    Some(Pn::Pn250)
                } else {
                    None
                }
            }
        },
        Pn::Pn320 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 250 {
                    Some(Pn::Pn320)
                } else {
                    None
                }
            }
        },
        Pn::Pn400 => match dn {
            Dn::Dn20 | Dn::Dn32 => None,
            _ => {
                if dn_val <= 200 {
                    Some(Pn::Pn400)
                } else {
                    None
                }
            }
        },
    }
}

#[inline(always)]
const fn get_pn_type_12(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn6 => {
            if dn_val <= 300 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 600 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 600 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        Pn::Pn63 => {
            if dn_val <= 40 {
                Some(Pn::Pn100)
            } else if dn_val <= 150 {
                Some(Pn::Pn63)
            } else {
                None
            }
        }
        Pn::Pn100 => {
            if dn_val <= 150 {
                Some(Pn::Pn100)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_13(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn6 => {
            if dn_val <= 300 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 600 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 150 {
                Some(Pn::Pn40)
            } else if dn_val <= 600 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 600 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        Pn::Pn63 => {
            if dn_val <= 40 {
                Some(Pn::Pn100)
            } else if dn_val <= 150 {
                Some(Pn::Pn63)
            } else {
                None
            }
        }
        Pn::Pn100 => {
            if dn_val <= 150 {
                Some(Pn::Pn100)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_35(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 => {
            if dn_val <= 1000 {
                Some(Pn::Pn2_5)
            } else {
                None
            }
        }
        Pn::Pn6 => {
            if dn_val <= 1200 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 1200 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 40 {
                Some(Pn::Pn40)
            } else if dn_val <= 1000 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        Pn::Pn25 => {
            if dn_val <= 125 {
                Some(Pn::Pn40)
            } else if dn_val <= 800 {
                Some(Pn::Pn25)
            } else {
                None
            }
        }
        Pn::Pn40 => {
            if dn_val <= 400 {
                Some(Pn::Pn40)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_36(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 | Pn::Pn6 => {
            if dn_val <= 500 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 => {
            if dn_val <= 150 {
                Some(Pn::Pn16)
            } else if dn_val <= 400 {
                Some(Pn::Pn10)
            } else {
                None
            }
        }
        Pn::Pn16 => {
            if dn_val <= 400 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[inline(always)]
const fn get_pn_type_37(pn: Pn, dn_val: u32) -> Option<Pn> {
    match pn {
        Pn::Pn2_5 | Pn::Pn6 => {
            if dn_val <= 200 {
                Some(Pn::Pn6)
            } else {
                None
            }
        }
        Pn::Pn10 | Pn::Pn16 => {
            if dn_val <= 200 {
                Some(Pn::Pn16)
            } else {
                None
            }
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Represents a single test case for the Synoptic Table logic.
    struct TestCase {
        flange_type: FlangeType,
        pn: Pn,
        dn: Dn,
        expected: Option<Pn>,
        description: &'static str,
    }

    impl TestCase {
        const fn new(
            flange_type: FlangeType,
            pn: Pn,
            dn: Dn,
            expected: Option<Pn>,
            description: &'static str,
        ) -> Self {
            Self {
                flange_type,
                pn,
                dn,
                expected,
                description,
            }
        }
    }

    /// Provides a comprehensive set of test cases covering all logic branches.
    /// Provides a comprehensive set of test cases covering all logic branches.
    fn get_test_cases() -> Vec<TestCase> {
        vec![
            // Type 01 - Plate flange for welding
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn2_5,
                Dn::Dn1000,
                Some(Pn::Pn6),
                "Type 01, PN 2.5, DN 1000 -> PN 6",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn2_5,
                Dn::Dn1200,
                Some(Pn::Pn2_5),
                "Type 01, PN 2.5, DN 1200 -> PN 2.5",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn2_5,
                Dn::Dn1400,
                None,
                "Type 01, PN 2.5, DN 1400 -> None",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn10,
                Dn::Dn40,
                Some(Pn::Pn40),
                "Type 01, PN 10, DN 40 -> PN 40",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn10,
                Dn::Dn150,
                Some(Pn::Pn16),
                "Type 01, PN 10, DN 150 -> PN 16",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn10,
                Dn::Dn600,
                Some(Pn::Pn10),
                "Type 01, PN 10, DN 600 -> PN 10",
            ),
            TestCase::new(
                FlangeType::Type01,
                Pn::Pn10,
                Dn::Dn800,
                Some(Pn::Pn10),
                "Type 01, PN 10, DN 800 -> PN 10",
            ),
            // Type 05 - High Pressure Exclusions
            TestCase::new(
                FlangeType::Type05,
                Pn::Pn160,
                Dn::Dn20,
                None,
                "Type 05, PN 160, DN 20 -> Excluded",
            ),
            TestCase::new(
                FlangeType::Type05,
                Pn::Pn160,
                Dn::Dn32,
                None,
                "Type 05, PN 160, DN 32 -> Excluded",
            ),
            TestCase::new(
                FlangeType::Type05,
                Pn::Pn160,
                Dn::Dn150,
                Some(Pn::Pn160),
                "Type 05, PN 160, DN 150 -> PN 160",
            ),
            TestCase::new(
                FlangeType::Type05,
                Pn::Pn400,
                Dn::Dn20,
                None,
                "Type 05, PN 400, DN 20 -> Excluded",
            ),
            TestCase::new(
                FlangeType::Type05,
                Pn::Pn400,
                Dn::Dn100,
                Some(Pn::Pn400),
                "Type 05, PN 400, DN 100 -> PN 400",
            ),
            // Type 32 - Weld-on plate collar
            TestCase::new(
                FlangeType::Type32,
                Pn::Pn2_5,
                Dn::Dn600,
                Some(Pn::Pn6),
                "Type 32, PN 2.5, DN 600 -> PN 6",
            ),
            TestCase::new(
                FlangeType::Type32,
                Pn::Pn2_5,
                Dn::Dn800,
                None,
                "Type 32, PN 2.5, DN 800 -> None",
            ),
            TestCase::new(
                FlangeType::Type32,
                Pn::Pn10,
                Dn::Dn40,
                Some(Pn::Pn40),
                "Type 32, PN 10, DN 40 -> PN 40",
            ),
            TestCase::new(
                FlangeType::Type32,
                Pn::Pn10,
                Dn::Dn150,
                Some(Pn::Pn16),
                "Type 32, PN 10, DN 150 -> PN 16",
            ),
            // Type 33 - Lapped pipe end (Identical to Type 37)
            TestCase::new(
                FlangeType::Type33,
                Pn::Pn2_5,
                Dn::Dn200,
                Some(Pn::Pn6),
                "Type 33, PN 2.5, DN 200 -> PN 6",
            ),
            TestCase::new(
                FlangeType::Type33,
                Pn::Pn2_5,
                Dn::Dn300,
                None,
                "Type 33, PN 2.5, DN 300 -> None",
            ),
            // Type 35 - Weldring neck
            TestCase::new(
                FlangeType::Type35,
                Pn::Pn2_5,
                Dn::Dn1000,
                Some(Pn::Pn2_5),
                "Type 35, PN 2.5, DN 1000 -> PN 2.5",
            ),
            TestCase::new(
                FlangeType::Type35,
                Pn::Pn25,
                Dn::Dn125,
                Some(Pn::Pn40),
                "Type 35, PN 25, DN 125 -> PN 40",
            ),
            TestCase::new(
                FlangeType::Type35,
                Pn::Pn25,
                Dn::Dn800,
                Some(Pn::Pn25),
                "Type 35, PN 25, DN 800 -> PN 25",
            ),
        ]
    }

    #[test]
    fn test_synoptic_table_compliance() {
        let cases = get_test_cases();
        let mut failures = Vec::new();

        for case in cases {
            let actual = get_target_pn(case.flange_type, case.pn, case.dn);
            if actual != case.expected {
                failures.push(format!(
                    "FAILED: {}\n  Input:    {:?}, {:?}, {:?}\n  Expected: {:?}\n  Actual:   {:?}",
                    case.description, case.flange_type, case.pn, case.dn, case.expected, actual
                ));
            }
        }

        if !failures.is_empty() {
            panic!(
                "Synoptic Table compliance check failed with {} errors:\n\n{}",
                failures.len(),
                failures.join("\n\n")
            );
        }
    }
}

*/
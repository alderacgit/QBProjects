// FileMode enum for QuickBooks session modes
#[derive(Debug, Clone, Copy)]
pub enum FileMode {
    SingleUser,
    MultiUser,
    DoNotCare,
    Online,
}

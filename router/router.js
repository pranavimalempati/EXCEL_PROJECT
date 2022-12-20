const controller = require("../controller/controller");
const {upload} = require ('../middleware/middleware');
const router = require("express").Router();

router.post('/uploads', upload.single('excel'),controller.uploadfile)

module.exports = router;
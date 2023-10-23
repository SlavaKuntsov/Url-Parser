import cheerio from 'cheerio'
import xl from 'excel4node'
import fs from 'fs'
import needle from 'needle'

const fetchData = async () => {
	const fileName = 'DataExcel.xlsx'

	// const workbook = new ExcelJS.Workbook()
	const workbook = new xl.Workbook()
	let worksheet
	// const worksheet = workbook.addWorksheet('База данных')

	let allArray = []
	let allImage = []

	let requests = []
	let rowCount // 1 with Headers

	let valuesRowObj = []
	let valuesRowArray = []

	const SAVE_INTERVAL = 1

	const headers = [
		'Дата прекращения действия',
		'Номер регистрации',
		'Дата регистрации',
		'Дата истечения срока регистрации',
		'Дата публикации',
		'Номер заявки',
		'Дата подачи заявки',
		'Приоритетная заявка',
		'Владелец',
		'Товары и/или услуги',
		'Изображение товарного знака'
	]
	//
	const fileExists = fs.existsSync(fileName)

	if (fileExists) {
		// await workbook.xlsx.readFile(fileName)
		worksheet = workbook.addWorksheet('База данных')
	} else {
		worksheet = workbook.addWorksheet('База данных')

		// worksheet.row(1).string(headers)
		headers.forEach((header, index) => {
			worksheet.cell(1, index + 1).string(header)
		})
		
	}

	rowCount = worksheet.rowCount // 1 with Headers
	// rowCount = 8401
	let counter = 0
	console.log('rowCount: ', rowCount)

	let ending = rowCount + 9 // 10
	let ending2 = rowCount + 99 // 100

	const ALL_ELEMENTS = 111394
	//

	const fetchItem = async function (i) {
		console.log(i)
		if (i > ALL_ELEMENTS) {
			// Выход из рекурсии
			return
		}

		// начало стало

		let imageUrl = `http://search.ncip.by/database/getimage.php?x=300&pref=tz&image=${i}`
		let imageBuffer

		let imageFilename = `assets/image${i}.jpeg`

		const imageExists = fs.existsSync(imageFilename)

		if (!imageExists) {
			const imageResponse = await needle('get', imageUrl)
			if (imageResponse.statusCode !== 200) {
				console.error(`Ошибка получения изображения для цели ${i}`)
				// imageBuffer = i
			} else {
				imageBuffer = imageResponse.body
				fs.writeFileSync(imageFilename, imageBuffer)
			}
		}

		allImage.push(imageBuffer)
		//
		const request = await needle(
			'get',
			`http://search.ncip.by/database/index.php?pref=tz&lng=ru&page=3&target=${i}`
		)
			.then(async function (response) {
				if (response.statusCode === 200) {
					const html = response.body
					const $ = cheerio.load(html)
					// console.log(html)

					const allNames = $('.col-md-6')

					let smallArray = []

					$(allNames).each(function (i, name) {
						const item = $(name).text()
						smallArray[i] = item
					})

					const result = []

					for (let j = 0; j < smallArray.length; j += 2) {
						const name = smallArray[j].replace(/[:\s]+$/, '')
						const value = smallArray[j + 1].trim()

						if (name !== 'Изображение товарного знака') {
							result.push({ name: name, value: value })
						}
					}

					// console.log('result: ', result);
					allArray.push(result)
				} else {
					console.error(`Ошибка запроса для цели ${i}`)
				}
			})
			.catch(function (err) {
				console.log('request')
				console.error(`Ошибка запроса для цели ${i}`)
				console.error(err)
			})
		//

		// конец
		counter++

		if (i % SAVE_INTERVAL === 0) {
			let itemCounter = 1
			let rowImageId =
				rowCount + itemCounter - 1 + counter - SAVE_INTERVAL
			// console.log('rowImageId: ', rowImageId)

			const fetchItemData = async function (j) {
				if (itemCounter === SAVE_INTERVAL + 1) {
					return
				}

				let imagePath = `assets/image${rowImageId}.jpeg`

				fs.access(imagePath, fs.constants.F_OK | fs.constants.R_OK, (err) => {
					if (err) {
						console.error('Ошибка чтения изображения:', err);
						return;
					}
					
					worksheet.addImage({
						path: imagePath,
						type: 'picture',
						position: {
							type: 'twoCellAnchor',
							from: { col: 11, colOff: 0, row: rowImageId + 1, rowOff: 0 },
							to: { col: 12, colOff: 0, row: rowImageId + 2, rowOff: 0 }
						}
					})
				})


				for (const header of headers) {
					const found = allArray[itemCounter - 1].find(
						obj => obj.name === header
					)
					if (found) {
						valuesRowObj.push(found.value)
					} else {
						valuesRowObj.push('null')
					}
					// console.log('found.value: ', found.value);
				}
				// const row = worksheet.addRow(valuesRowObj)
				// row.height = 200
				valuesRowObj.forEach((value, index) => {
					if(index + 1 === 11) {
						worksheet.cell(rowImageId + 1, 11).string('image')
					}
					else{
						worksheet.cell(rowImageId + 1, index + 1).string(value)
					}
				})

				worksheet.row(rowImageId).setHeight(50);

				valuesRowArray.push(valuesRowObj)
				valuesRowObj = []

				// let rowImageId = rowCount    + itemCounter - 1      + counter - SAVE_INTERVAL     - rowCount + 1
				// console.log('rowCount: ', rowCount);
				// console.log('itemCounter: ', itemCounter);
				// console.log('counter: ', counter);
				// console.log('SAVE_INTERVAL: ', SAVE_INTERVAL);
				// console.log('--------- rowImageId: ', rowImageId);


				// if (allImage.includes(rowImageId)) {
				// 	// const imageId = workbook.addImage({
				// 	// 	filename: `assets/cake.jpeg`,
				// 	// 	extension: 'jpeg'
				// 	// })
				// 	// console.log('imageId: err', imageId);
				// 	// worksheet.addImage(imageId, {
				// 	// 	tl: { col: 10, row: rowImageId },
				// 	// 	ext: { width: 200, height: 200 }
				// 	// })
				// } else {
				// 	const imageId = workbook.addImage({
				// 		filename: `assets/image${rowImageId}.jpeg`,
				// 		extension: 'jpeg'
				// 	})
				// 	// }) + 8401
				// 	console.log('imageId: ', imageId)
				// 	// console.log('filename: assets/image${rowImageId}.jpeg: ', `filename: assets/image${rowImageId}.jpeg`);
				// 	// console.log('imageId: okk', imageId);

				// 	// let img = rowImageId + rowCount

				// 	worksheet.addImage(imageId, {
				// 		tl: { col: 10, row: rowImageId },
				// 		ext: { width: 200, height: 200 }
				// 	})
				// 	// console.log('add')
				// }
				valuesRowArray = []

				itemCounter++
				await fetchItemData(j + 1)
			}

			try {
				// console.log('fetchItemData i ' + i)
				await fetchItemData(i) // приходит i = 10
				worksheet.row(1).setHeight(15)
				await saveData(workbook, fileName, i, SAVE_INTERVAL)
				allArray = []
			} catch (err) {
				console.log('fetchItemData')
				console.error(err)
			}
		}

		// console.log(i)
		requests.push(request)

		await new Promise(resolve => setTimeout(resolve, 300))

		await fetchItem(i + 1)
	}
	/////////////////////////////////////////////////////////////////
	try {
		await fetchItem(rowCount)
	} catch (err) {
		console.log('fetchItem')
		console.error(err)
		process.exit(1)
	}
	/////////////////////////////////////////////////////////////////
	Promise.all(requests)
		.then(() => {
			// worksheet.addRow(headers)
			// let valuesRowObj1 = []
			// let valuesRowArray1 = []
			// let rowCount = worksheet.rowCount
			// console.log('rowCount in promise: ', rowCount)
			// console.log(allArray)
			// // create
			// allArray.forEach(array => {
			// 	for (const header of headers) {
			// 		const found = array.find(obj => obj.name === header)
			// 		if (found) {
			// 			valuesRowObj1.push(found.value)
			// 		} else {
			// 			valuesRowObj1.push('null')
			// 		}
			// 	}
			// 	valuesRowArray1.push(valuesRowObj1)
			// 	valuesRowObj1 = []
			// })
			// // add
			// valuesRowArray1.forEach((array, i) => {
			// 	const row = worksheet.addRow(array)
			// 	row.height = 200
			// 	let rowImageId = rowCount + i
			// 	if(allImage.includes(rowImageId)) {
			// 	}
			// 	else{
			// 		const imageId = workbook.addImage({
			// 			filename: `assets/image${rowImageId}.jpeg`,
			// 			extension: 'jpeg'
			// 		})
			// 		worksheet.addImage(imageId, {
			// 			tl: { col: 10, row: rowImageId },
			// 			ext: { width: 200, height: 200 }
			// 		})
			// 	}
			// })
			// push
			// saveData(workbook, fileName, 'end')
		})
		.catch(error => {
			console.log('end')
			console.error(error)
		})
}

fetchData()

const saveData = async (workbook, fileName, i, saveInternal) => {
	try {

		if (i !== saveInternal) {
			createBackup(fileName)
		}
		await workbook.write(fileName)
		// console.log(`Файл сохранен. Текущая итерация ${i}`)
	} catch (error) {
		console.error(`Ошибка при сохранении данных save:`, error)
		process.exit(1) // Завершение выполнения программы
	}
}

const createBackup = async fileName => {
	try {
		const backupFileName = `${fileName}.bak`
		fs.copyFileSync(fileName, backupFileName)
		// console.log(`Создан бэкап файла: ${backupFileName}`)
	} catch (error) {
		console.error('Ошибка при создании бэкапа:', error)
		process.exit(1) // Завершение выполнения программы
	}
}

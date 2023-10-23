import cheerio from 'cheerio'
import ExcelJS from 'exceljs'
import fs from 'fs'
import needle from 'needle'

const fetchData = async () => {
	let fileName = 'Data.xlsx'

	const workbook = new ExcelJS.Workbook()
	let worksheet
	// const worksheet = workbook.addWorksheet('База данных')

	let allArray = []
	let allImage = []

	let requests = []
	let rowCount // 1 with Headers

	const SAVE_INTERVAL = 10
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
	const fileExists = fs.existsSync('Data.xlsx')

	if (fileExists) {
		await workbook.xlsx.readFile('Data.xlsx')
		worksheet = workbook.getWorksheet('База данных')
	} else {
		worksheet = workbook.addWorksheet('База данных')

		worksheet.addRow(headers)
	}

	rowCount = worksheet.rowCount // 1 with Headers
	console.log('rowCount: ', rowCount)

	let ending = rowCount + 9 // 10
	let ending2 = rowCount + 99 // 100

	const ALL_ELEMENTS = 111394
	//
	try {
		for (let i = rowCount; i <= ending2; i++) {
			let imageUrl = `http://search.ncip.by/database/getimage.php?x=300&pref=tz&image=${i}`
			let imageBuffer

			const imageResponse = await needle('get', imageUrl)
			if (imageResponse.statusCode !== 200) {
				console.error(`Ошибка получения изображения для цели ${i}`)
				imageBuffer = i
			}
			else{
				imageBuffer = imageResponse.body
				fs.writeFileSync(`assets/image${i}.jpeg`, imageBuffer)
			}

			requests.push(imageResponse)

			allImage.push(imageBuffer)

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

						allArray.push(result)


					} else {
						console.error(`Ошибка запроса для цели ${i}`)
					}
				})
				.catch(function (err) {
					console.log('request')
					console.log(err)
				})



			// if (i % SAVE_INTERVAL === 0) {
				// await new Promise(resolve => {
					// resolve(saveData(workbook, 'Data.xlsx', i))
					await saveData(workbook, 'Data.xlsx', i)
					
				// }).then(() => {
				// 	console.log('dddddddddddd')

				// })
			// }
			requests.push(request)
			console.log(i)
			await new Promise(resolve => setTimeout(resolve, 200))
		}
		// await saveData(workbook, 'Data.xlsx', 'end')
	} catch (err) {
		console.log('for')
		console.log(err)
	}

	Promise.all(requests)
		.then(() => {
			// worksheet.addRow(headers)

			let valuesRowObj = []
			let valuesRowArray = []
			let rowCount = worksheet.rowCount
			console.log('rowCount in promise: ', rowCount)

			// console.log(array)

			// create
			allArray.forEach(array => {
				for (const header of headers) {
					const found = array.find(obj => obj.name === header)
					if (found) {
						valuesRowObj.push(found.value)
					} else {
						valuesRowObj.push('null')
					}
				}
				valuesRowArray.push(valuesRowObj)
				valuesRowObj = []
			})

			// add
			valuesRowArray.forEach((array, i) => {
				const row = worksheet.addRow(array)
				row.height = 200

				let rowImageId = rowCount + i

				if(allImage.includes(rowImageId)) {

				}
				else{
					const imageId = workbook.addImage({
						filename: `assets/image${rowImageId}.jpeg`,
						extension: 'jpeg'
					})
	
					worksheet.addImage(imageId, {
						tl: { col: 10, row: rowImageId },
						ext: { width: 200, height: 200 }
					})
				}

			})

			// push
			saveData(workbook, 'Data.xlsx', 'end')
		})
		.catch(error => {
			console.log('end')
			console.error(error)
		})
}

fetchData()

const saveData = async (workbook, fileName, i) => {
	try {
		await workbook.xlsx.writeFile(fileName)
		console.log(`Файл сохранен. Текущая итерация ${i}`)

		if(i !== 10) {
			createBackup(fileName)
		}

	} catch (error) {
		console.error('Ошибка при сохранении данных:', error)
	}
}

const createBackup = async (fileName) => {
	try {
		const backupFileName = `${fileName}.bak`
		fs.copyFileSync(fileName, backupFileName)
		console.log(`Создан бэкап файла: ${backupFileName}`)
	} catch (error) {
		console.error('Ошибка при создании бэкапа:', error)
	}
}

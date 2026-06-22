from kafka import KafkaProducer

producer = KafkaProducer(
    bootstrap_servers='localhost:9092',
    value_serializer=lambda v: v.encode('utf-8')
)

while True:
    producer.send('etl-commands', input("Команда: "))
    producer.flush()
    print("Отправлено")
from kafka import KafkaConsumer

consumer = KafkaConsumer(
    'etl-results',
    bootstrap_servers='localhost:9092',
    value_deserializer=lambda m: m.decode('utf-8'),
    auto_offset_reset='earliest',
    group_id='result-reader'
)

print("Жду результат...")
for message in consumer:
    print(f"Получено: {message.value}")